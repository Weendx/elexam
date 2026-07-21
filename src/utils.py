import datetime
from typing import Union, List, Optional, Iterable
import secrets
import string
import pyperclip
import subprocess

from datatypes import (
    UserInfo, UserAction, 
    UserActionType, Course, UserTableData
)

from label import LabelController, LabelControllerError

def pluralize(number, forms):
    """
    Возвращает правильную форму слова в зависимости от числа.

    Args:
        number: Число, для которого нужно склонять слово.
        forms: Список из трех форм слова:
               [единственное число, множественное число для 2-4, множественное число для остальных]
               Например: ['яблоко', 'яблока', 'яблок']

    Returns:
        Строка: Слово в правильной форме.
    """
    n = abs(number)
    if n % 10 == 1 and n % 100 != 11:
        return forms[0]
    elif 2 <= n % 10 <= 4 and (n % 100 < 10 or n % 100 >= 20):
        return forms[1]
    else:
        return forms[2]

def convert_date_string(datetime_string: str) -> Union[datetime.datetime, None]:
        if not datetime_string: return None
        
        months = ["июнь", "июль", "август"]
        if datetime_string in months:
            month = months.index(datetime_string) + 6
            return datetime.datetime(2001, month, 1)

        formats = ["%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"]
        for fmt in formats:
            try:
                return datetime.datetime.strptime(datetime_string, fmt)
            except ValueError:
                pass
        return None

def convert_block_dates(dates: Iterable[Union[datetime.date, datetime.datetime]]) -> str:
    """ Convert block datetime dates to string like '21.06,21.07' """
    return ",".join([(x.strftime('%d.%m') if x != datetime.date(2000,1,1) else '-') for x in dates])

def generate_random_string(length=6) -> str:
    """Генерирует случайную строку заданной длины, используя безопасный генератор."""
    alphabet = string.ascii_letters + string.digits  # Буквы (верхний и нижний регистр) + цифры
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def get_mock_user() -> UserInfo:
    courses = list()
    courses.append(Course(
        1,
        "Вступительный курс 1", 
        datetime.datetime.strptime("20.11.2024", "%d.%m.%Y"), 
        datetime.datetime.strptime("11.01.2025", "%d.%m.%Y"),
        ("Исаев Валерий Васильевич", "Лукьянова Юлия Михайловна")
    ))
    courses.append(Course(
        1,
        "Вступительный экзамен 2", 
        datetime.datetime.strptime("05.07.2024", "%d.%m.%Y"), 
        datetime.datetime.strptime("11.01.2025", "%d.%m.%Y")
    ))
    registered = datetime.datetime.strptime("31.08.2024", "%d.%m.%Y")
    logdate = datetime.datetime.strptime("28.06.2025 21:16:40", "%d.%m.%Y %H:%M:%S")

    return UserInfo(
        mid=49460, 
        login="25-01645",
        email="test@test.test",
        fio="Жмышенко Валерий Test 002", 
        registered=registered,
        table=UserTableData("test@test.test", "25-01645"),
        last_login=logdate,
        courses=tuple(courses)
    )

def generate_password(user_table_login: str = ""):
    if not user_table_login:
        return 'EL_' + generate_random_string(6)
    try:
        return (int(user_table_login[-5:])+23000)*15
    except (ValueError):
        return 'EL_' + generate_random_string(6)

def suggest_user_actions(uinfo: UserInfo, learning = None) -> List[UserAction]:
    suggestions = []
    if uinfo.registered:
        current_date = datetime.datetime.today()
        last_autumn = datetime.datetime(year=current_date.year - 1, month=11, day=1)
        if uinfo.registered < last_autumn:
            suggestions.append(UserAction.DELETE)
            return suggestions

    if uinfo.table:
        if uinfo.table.email and uinfo.table.email.lower() != uinfo.email.lower():
            # возможно, пользователь найден по ошибке, т.к. используется нечеткий поиск
            # поэтому пропускаем
            suggestions.append(UserAction.SILENT_SKIP)
            return suggestions

        if uinfo.login and uinfo.table.login:
            if uinfo.login != uinfo.table.login:
                suggestions.append(UserAction.CHANGE_LOGIN)

                # Смена пароля
                if uinfo.source != "AD":
                    password = "<Неизвестно>"
                    
                    # Пробуем узнать пароль
                    if learning:
                        try:
                            password = learning.get_user_password(uinfo.mid)
                        except Exception:
                            pass

                    if password != "<Неизвестно>":
                        suggestions.append(UserAction(UserActionType.CHANGE_PASSW_LOCAL, password))
                    elif uinfo.tags and any(('spo' in x for x in uinfo.tags)) or \
                        (uinfo.table.subjects and any( 
                                ('подготовка к егэ' in x.name.lower() or \
                                        'на базе спо' in x.name.lower() \
                                            for x in uinfo.table.subjects
                                ) 
                            )
                        ):
                        suggestions.append(UserAction(UserActionType.CHANGE_PASSW_LOCAL, password))
                    else:
                        password = generate_password(uinfo.table.login)
                        suggestions.append(UserAction(UserActionType.CHANGE_PASSW_EDU, password))
                        suggestions.append(UserAction(UserActionType.CHANGE_PASSW_LOCAL, password))
                ##
        ##

        if uinfo.table.subjects:
            appended = False
            to_delete = []

            for subject in uinfo.table.subjects:
                try:
                    # label = str(LabelController.get_label_primitive(subject))
                    selected_date = convert_date_string(str(subject.date)).date() if subject.date else None
                    label = LabelController.get_label(subject.name, selected_date=selected_date)

                    if not uinfo.tags or (uinfo.tags and label not in uinfo.tags):
                        suggestions.append(UserAction(UserActionType.ADD_LABEL, label))
                        appended = True
                    elif not (uinfo.table.marks and 'registered' in uinfo.table.marks):
                        to_delete.append(UserAction(UserActionType.DELETE_FROM_TABLE_WITH_SUBJECT, subject.name))

                except LabelControllerError:
                    print("\n\033[31mSUGGESTION ERROR: Cannot get label for subject", subject, "\033[0m")
            
            if len(to_delete) == len(uinfo.table.subjects):
                suggestions.append(UserAction(UserActionType.DELETE_FROM_TABLE))
            elif to_delete: 
                for suggestion in to_delete:
                    suggestions.append(suggestion)
                appended = True

            if appended:
                suggestions.append(UserAction(UserActionType.MARK_REGISTERED))
    
    return suggestions

def copy_to_clipboard(text):
    try:
        # Пытаемся скопировать через обычный pyperclip (для других платформ)
        pyperclip.copy(text)
    except pyperclip.PyperclipException:
        try:
            # Используем termux-api
            subprocess.run(["termux-clipboard-set", text],
                       stdin=subprocess.DEVNULL,
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL,
                       check=True)
        except FileNotFoundError:
            print("termux-api не установлен.  Выполните: pkg install termux-api")
        except subprocess.CalledProcessError as e:
            print(f"Ошибка при копировании в буфер обмена Termux: {e}")

def is_blue_color(hex_color: str) -> bool:
    if hex_color == 'FF558ED5': return True
    if len(hex_color) == 8:  # AARRGGBB
        hex_color = hex_color[2:]
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)

    return blue > red and blue > green

def is_red_color(hex_color: str) -> bool:
    if len(hex_color) == 8:  # AARRGGBB
        hex_color = hex_color[2:]
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)

    min_delta = 125

    return red > blue and red > green and red - blue > min_delta and red - green > min_delta

def is_array_consecutive(array) -> bool:
    """ Проверка, что list/tuple содержит последовательные элементы : [1,2,3] """
    if not array:
        return True
    if len(array) == 1:
        return True
    a = sorted(array)
    return all(a[i] + 1 == a[i+1] for i in range(len(a)-1))

