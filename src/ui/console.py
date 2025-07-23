import os
import re
import sys
from sys import exit
import datetime
import traceback
from typing import List, Tuple, Optional, Callable, Iterable

from rich import print
from rich.console import Console as RichConsole
from rich.console import Group, RenderableType
from rich.panel import Panel
from rich.prompt import Prompt, IntPrompt, Confirm
from rich.progress import track
from rich.table import Table
from rich.text import Text
from rich.live import Live

sys.path.append("..")
from settings import Settings
from label import LabelController, NoSuitableLabelFound
from learning import (
    LearningDriver,
    InvalidLoginPair, DataAgreementNotAccepted, 
    RequestError, UserNotFound, NotAuthorized
)
from excelDriver import ExcelDriver, UserTableData
from excelDriver import UserNotFoundException as ExcelUserNotFound
from fileController import FileController
from datatypes import (
    MenuItem, UserInfo, UserAction, 
    UserActionType, AuthCookies, Exam
)
from utils import (
    pluralize, get_mock_user, 
    suggest_user_actions, 
    copy_to_clipboard, generate_password
)


class Console(RichConsole):

    AUTHCOOKIEID = "auth"

    def __init__(self):
        super().__init__()
        self.parse_args()

    def _create_menu(self):
        """ Returns menu renderable, menu items and menu length for run_menu() """
        menuitems = list()
        logged = bool(Settings()[self.AUTHCOOKIEID])
        if logged:
            menuitems.append(MenuItem('Проверить вход в систему', 'auth_check'))
            menuitems.append(MenuItem('Выполнить обработку файла (часть 1)', 'process_file_1'))
        else:
            menuitems.append(MenuItem('Войти в систему', 'auth_check'))
        menuitems.append(MenuItem('Выполнить обработку файла (часть 2, csv)', 'process_file_2'))
        if logged:
            menuitems.append(MenuItem('Выполнить действия с пользователем (eLearning)', 'perform_actions'))
            menuitems.append(MenuItem('Получить информацию о пользователе (eLearning)', 'show_uinfo'))
            menuitems.append(MenuItem('Найти пароль пользователя (eLearning)', 'find_password'))
        menuitems.append(MenuItem('Управление метками', 'manage_labels'))
        if logged:
            menuitems.append(MenuItem('Выйти из eLearning', 'logout'))
        menuitems.append(MenuItem('Выход', 'exit'))
        menuitems_len = len(menuitems)
        menutext = " " + "\n ".join([f"{x + 1}. {y}" for x, y in enumerate(map(lambda x: x.title, menuitems))])
        menu = Panel.fit(menutext, title="= ==== ==  [italic]Меню[/italic]  == ==== =", border_style="yellow")
        return menu, menuitems, menuitems_len
    
    def _create_labels_menu(self):
        """ Returns menu renderable, menu items and menu length for run_labels_menu() """
        menuitems = list()
        menuitems.append(MenuItem('Добавить экзамен', 'add'))
        menuitems.append(MenuItem('Просмотр экзаменов', 'show_all'))
        menuitems.append(MenuItem('Редактировать экзамен', 'edit'))
        menuitems.append(MenuItem('Удалить экзамен', 'delete'))
        menuitems.append(MenuItem('Определить метку', 'test'))
        menuitems.append(MenuItem('Строка обмена', 'share'))
        menuitems.append(MenuItem('Назад', 'exit'))
        menuitems_len = len(menuitems)
        menutext = " " + "\n ".join([f"{x + 1}. {y}" for x, y in enumerate(map(lambda x: x.title, menuitems))])
        menu = Panel.fit(menutext, title="= ==== ==  [italic]Управление метками[/italic]  == ==== =", border_style="yellow")
        return menu, menuitems, menuitems_len

    @staticmethod
    def ask(message: str, 
            validator: Optional[Callable[[str], bool]] = None, default=None) -> str:
        value = None
        while not value or (validator and not validator(value)):
            try:
                value = Prompt.ask(message, default=default)
            except KeyboardInterrupt:
                print()
        return value
    
    @staticmethod
    def ask_int(message: str, min_ = None, max_ = None,
                    validator: Optional[Callable[[int], bool]] = None, default=None) -> int:
        value = None
        while (value == None or (min_ and value < min_) or \
            (max_ and value > max_) or (validator and not validator(value))):
            try:
                value = IntPrompt.ask(message, default=default)
            except KeyboardInterrupt:
                print()
        return value

    def ask_exam_number(self, exams: Iterable[Exam]) -> int:
        self.print(self.compose_exams_table(exams))
        return self.ask_int("Введите номер", 1, len(exams))

    def ask_exam_params(self, exam: Optional[Exam] = None) -> Exam:
        subject = exam.subject if exam else None
        tag = exam.tag if exam else None
        dates = ",".join([(x.strftime('%d.%m') if x != datetime.date(2000,1,1) else '-') for x in exam.dates]) if exam else None

        subject = self.ask("[cyan]Введите название[/cyan] [dim](предмет)[/dim]", default=subject)
        tag = self.ask("[cyan]Введите метку[/cyan] [dim](без года и номера блока)[/dim]", default=tag)
        
        new_dates = list()
        self.print("\n[cyan]Введите даты проведения экзамена:")
        self.print("  [dim]ДД.ММ, через запятую без пробела.")
        self.print("  [dim]Номер блока будет определяться в соответствии порядку записи дат.")
        self.print("  [dim]Если в блоке экзамена нет, введите «-»")
        self.print(Text("  Пример: 18.06,-,20.08", style="dim"))
        reexp = re.compile('^((([0-9]{2}\\.[0-9]{2})|(-)),{0,1})+$')
        while True:
            _dates = self.ask("[cyan]Ввод", default=dates)
            if not reexp.match(_dates): continue
            try:
                input_dates = _dates.split(',')
                for input_date in input_dates:
                    if input_date == '-':
                        new_dates.append(datetime.date(2000, 1, 1))
                        continue
                    date_parts = input_date.split('.')
                    new_dates.append(datetime.date(year=2000, month=int(date_parts[1]), day=int(date_parts[0])))
                break
            except:
                continuraisee

        return Exam(subject, tag, new_dates)

    def ask_filepath(self, required=True, cache_suffix='filepath') -> str:
        """ Returns path of the file
            required(bool=True): you can abort asking by ! when required is False
            cache_suffix(str='filepath'): keyvalue to store input in cache
        """
        settings = Settings()
        filepath = ""
        is_first = True
        while not os.path.isfile(filepath):
            message = "Введите путь к файлу"
            if not is_first: message = "Введите корректный путь к файлу"
            filepath = Prompt.ask(message, default=settings[f'cache.{cache_suffix}'])
            if '~' in filepath:
                filepath = os.path.expanduser(filepath)
            is_first = False
            if not required and filepath == '!':
                return filepath
        settings[f'cache.{cache_suffix}'] = filepath
        return filepath

    def ask_hmkey(self) -> str:
        """ Returns hmkey cookie value """
        return Prompt.ask("Enter your hmkey cookie")

    def ask_login(self) -> str:
        """ Returns hmkey cookie value after log in the system """
        login = Prompt.ask("Enter your eLearning login")
        password = Prompt.ask("Enter your eLearning password", password=True)

    def ask_user_action(self, abilities: Iterable, with_help=False,) -> str:
        excel = 'excel' in abilities
        learning = 'learning' in abilities
        choices = []
        helptext = ""

        items = [
            {'sign': '!', 'desc': 'Перейти к след. пользователю', 'requires': (True)},
            {'sign': '!!', 'desc': 'Выйти', 'requires': (True)},
            {'sign': '?', 'desc': 'Показать справку', 'requires': (True)},
            {'sign': 'c', 'desc': 'Установить комментарий в таблице', 'requires': (excel)},
            {'sign': 'd', 'desc': 'Удалить пользователя', 'requires': (learning)},
            {'sign': 'dt', 'desc': 'Удалить пользователя из таблицы', 'requires': (excel)},
            {'sign': 'k', 'desc': 'Пропустить без пометки в таблице', 'requires': (True)},
            {'sign': 'l', 'desc': 'Обновить логин в таблице', 'requires': (learning, excel)},
            {'sign': 'm', 'desc': 'Назначить метку', 'requires': (learning)},
            {'sign': 'md', 'desc': 'Убрать метку', 'requires': (learning)},
            {'sign': 'pe', 'desc': 'Установить пароль в eLearning', 'requires': (learning)},
            {'sign': 'pg', 'desc': 'Сгенерировать пароль', 'requires': (True)},
            {'sign': 'pl', 'desc': 'Установить пароль в таблице', 'requires': (excel)},
            {'sign': 'r', 'desc': 'Очистить список действий', 'requires': (True)},
            {'sign': 'ro', 'desc': 'Удалить одно действие', 'requires': (True)},
            {'sign': 's', 'desc': 'Пропустить с пометкой в таблице', 'requires': (excel)},
            {'sign': 'sr', 'desc': 'Отметить в таблице, что зарегистрирован', 'requires': (excel)},
            {'sign': 'x', 'desc': 'Показать доп. информацию', 'requires': (True)},
        ]
        
        for item in items:
            if type(item['requires']) == bool:
                if not item['requires']:
                    continue
            elif not all(item['requires']):
                continue

            choices.append(item['sign'])
            if with_help:
                prefix = "\n " if helptext else " "
                helptext += prefix + f"[magenta]{item['sign']:2}[/magenta]- {item['desc']}"

        if with_help:
            self.print(Panel.fit(helptext))

        return Prompt.ask(
            "Выберите действие (? для справки)",
            choices=choices,
            case_sensitive=False, console=self
        )

    def compare_passwords(self, passw1, passw2, print_:bool=True) -> bool:
        """ print_(bool): show error output """
        if passw1 == passw2: return True
        if print_:
            self.print("[red]Err: [/red][dim][italic][not bold]Пароли не совпадают (таблица/eLearning)")
        return False

    @classmethod
    def compose_exams_table(cls, exams: Iterable[Exam]) -> RenderableType:
        if not exams: return "[blue]Список экзаменов пуст"

        table = Table()
        table.add_column("#")
        table.add_column("Предмет")
        table.add_column("Тег")
        table.add_column("Даты")

        for i, exam in enumerate(exams):
            dates = ",".join([(x.strftime('%d.%m') if x != datetime.date(2000,1,1) else '-') for x in exam.dates])
            table.add_row(str(i+1), exam.subject, exam.tag, dates, style="blue" if i%2==0 else 'cyan')
        return table

    def compose_users_actions(self, 
                users_actions: List[Tuple[UserInfo, List[UserAction]]]) -> RenderableType:
        text = ""
        
        emails = list()

        _is_first_run = True
        for row in users_actions:
            uinfo = row[0]
            gap = "\n\n" if not _is_first_run else ""
            fio = uinfo.table.fio if uinfo.table and uinfo.table.fio else uinfo.fio
            text += gap + f"[cyan]{fio} ({uinfo.login}, {uinfo.email})[/cyan]"
            emails.append(uinfo.email)
            for uaction in row[1]:
                text += f"\n[dim]-[/dim] {uaction!s}"
            _is_first_run = False

        # just a short hook
        emails_str = ", ".join(emails)
        Settings()['cache.file_step1_emails'] = emails_str

        return Group(
            Panel.fit(text, title="Выбранные действия", border_style="dim"), 
            Text("\nemails:", style="dim", end=" "), Text(emails_str)
        ) 

    def compose_user_info(self, user: UserInfo, detailed=True, title:str|None = None) -> RenderableType:
        courses_count = len(user.courses) if user.courses else 0
        courses_word = pluralize(courses_count, ("курс", "курса", "курсов"))

        courses = list()
        if user.courses:
            for course in user.courses:
                if not detailed:
                    coursetext = f"[dim]-[/dim] [blue]Курс {course.title}[/blue]"
                    courses.append(coursetext)
                    continue

                coursedate = ""
                if course.starts and course.ends:
                    coursedate = f"с {course.starts.strftime('%d.%m.%Y')} по {course.ends.strftime('%d.%m.%Y')}"
                elif course.starts:
                    coursedate = f"с {course.starts.strftime('%d.%m.%Y')}"
                elif course.ends:
                    coursedate = f"по {course.ends.strftime('%d.%m.%Y')}"

                coursetext = f"[cyan]Курс {course.title}[/cyan]"
                coursetext += f"\n[blue]ID курса:[/blue] {course.cid}"
                coursetext += f"\n[blue]Проводится:[/blue] {coursedate if coursedate else "-"}"
                coursetext += f"\n[blue]Тьюторы:[/blue] {', '.join(course.teachers) if course.teachers else "-"}"
                courses.append(coursetext)

            coursepanel = Panel("\n\n".join(courses) if detailed else "\n".join(courses), 
                title="Курсы пользователя", 
                # padding=(1,2), 
                border_style="dim white"
            )

        usertext_registered = "-"
        if user.registered:
            current_date = datetime.datetime.today()
            last_autumn = datetime.datetime(year=current_date.year - 1, month=9, day=1)
            s = 'green' if user.registered >= last_autumn else 'red'
            usertext_registered = f"[{s}]" + user.registered.strftime('%d.%m.%Y') + f"[/{s}]"
        
        usertext_last_login = "-"
        if user.last_login:
            current_date = datetime.datetime.today()
            s = 'green' if abs((user.last_login - current_date).days) > 15 else 'red'
            usertext_last_login = f"[{s}]" + user.last_login.strftime('%d.%m.%Y %H:%M:%S') + f"[/{s}]"

        usertext_login = "-"
        if user.login and (not user.table or not user.table.login):
            m = re.match("[0-9]{2,3}-[0-9]{5}", user.login)
            s = 'green' if m else 'red'
            usertext_login = f"[{s}]{user.login}[/{s}]"
        elif user.login and (user.table and user.table.login):
            m = user.login == user.table.login
            s = 'green' if m else 'red'
            o = '==' if m else '!='
            usertext_login = f"[{s}]{user.login} {o} {user.table.login}[/{s}]"

        usertext_source = "-"
        if user.source:
            colors = {'ELS': 'green', 'AD': 'bold red', 'elexam': 'italic blue'}
            color = colors.get(user.source)
            if color: usertext_source = f"[{color}]{user.source}[/{color}]"
            else: usertext_source = user.source

        usertext_subjects = "-"
        if user.table and user.table.subjects:
            tmp = []
            for subject in user.table.subjects:
                if subject.date:
                    date = subject.date.strftime('%d.%m.%Y')
                    tmp.append(f"{subject.name} <{date}>")
                else:
                    tmp.append(subject.name)
            usertext_subjects = '; '.join(tmp)

        userp = ' ' if courses_count else '' # usertext_padding
        usertext = ""
        usertext += f"{userp}[cyan]MID:[/cyan] {user.mid}\n" if detailed else ""
        usertext += f"{userp}[cyan]Email:[/cyan] {user.email}"
        usertext += f"\n{userp} [red]В таблице {user.table.email}[/red]" if user.table \
                 and user.table.email and user.table.email.lower() != user.email.lower() else ""
        usertext += f"\n{userp}[cyan]ФИО:[/cyan] {user.fio}"
        usertext += f"\n{userp} [red]В таблице {user.table.fio}[/red]" if user.table \
                 and user.table.fio and user.table.fio.lower() != user.fio.lower() else ""
        usertext += f"\n{userp}[cyan]Логин:[/cyan] {usertext_login}"
        usertext += f"\n{userp}[cyan]Зарегистрирован:[/cyan] {usertext_registered}"
        usertext += f"\n{userp}[cyan]Последний вход:[/cyan] {usertext_last_login}"
        usertext += f"\n{userp}[cyan]Создан в:[/cyan] {usertext_source}"
        usertext += f"\n{userp}[cyan]Кол-во курсов:[/cyan] {str(courses_count) + ' ' + courses_word if courses_count else '-'}"
        usertext += f"\n{userp}[cyan]Метки:[/cyan] {', '.join(user.tags) if user.tags else '-'}"
        usertext += f"\n{userp}[cyan]Предметы:[/cyan] {usertext_subjects}"
        if courses_count: usertext = "\n" + usertext + "\n"

        userpanel_renderable = Group(usertext, coursepanel) if courses_count else usertext 

        userpanel = Panel.fit(userpanel_renderable, 
            title=title if title else "Информация о пользователе", 
            # padding=(1,2), 
            border_style="dim cyan",
        )
        
        return userpanel

    def confirm_users_actions(self, 
                users_actions: List[Tuple[UserInfo, List[UserAction]]]) -> bool:
        ##
        selection = None

        while True:
            selection = Prompt.ask(
                "[yellow]Все действия выбраны. Продолжить?[/yellow] [bright_black](x - просмотр, c - copy emails)[/bright_black]",
                choices=['y', 'n', 'x', 'c'], case_sensitive=False
            )
            if selection == 'y': 
                self.elearning_auth_check(silent=True)
                return True
            elif selection == 'n': return False
            elif selection == 'x':
                self.print()
                self.print(self.compose_users_actions(users_actions))
                self.print()
            elif selection == 'c':
                emails = ", ".join([row[0].email for row in users_actions])
                try:
                    copy_to_clipboard(emails)
                    self.print("[green]Адреса email скопированы в буфер!")
                except:
                    self.print("[red]Не получается скопировать данные в буфер :(")

    @classmethod
    def create_learning(cls) -> LearningDriver:
        ac = None
        ac_list = Settings().get_crypted(cls.AUTHCOOKIEID)
        if ac_list:
            try:
                ac = AuthCookies(*ac_list)
            except:
                pass
        return LearningDriver(auth_cookies=ac)

    def elearning_auth_check(self, silent=False) -> None:
        """ menu action
            silent: show only important output
        """
        ##
        learning = self.create_learning()

        with self.status("Проверка входа... ") as status:
            try:
                is_logged = learning.auth_check()
            except DataAgreementNotAccepted:
                self.print("[red]Не принято согласие на обработку персональных данных![/red]")
                self.print("Войдите в систему в браузере и примите согласие на главной странице")
                exit()
            if is_logged:
                fio = learning.get_current_fio()
                role = learning.get_current_role()
                if silent == False:
                    self.print(f"[green]Вход выполнен под именем {fio} [not bold]({role})[/not bold][/green]")
                return
            else:
                self.print("[dim blue]Необходимо выполнить вход[/dim blue]")

        
        _is_first_run = True
        while not is_logged:
            if not _is_first_run: self.print()
            login = Prompt.ask("Введите логин")

            if _is_first_run: self.print("\n[dim]Пароль может не отображаться при вводе, это нормально[/dim]")
            passw = Prompt.ask("Введите пароль", password=True)

            try:
                with self.status("Вход в систему... ") as status:
                    ac = learning.auth(login, passw)
                    try:
                        is_logged = learning.auth_check()
                    except DataAgreementNotAccepted:
                        self.print("\n[red]Не принято согласие на обработку персональных данных![/red]")
                        self.print("Войдите в систему в браузере и примите согласие на главной странице")
                        exit()
                    if is_logged:
                        Settings().set_crypted(self.AUTHCOOKIEID, ac)
                        self.print("\n[green]Аутентификация прошла успешно![/green]")
            except InvalidLoginPair:
                self.print("\n[red]Неверный логин или пароль[/red]")

            _is_first_run = False
        ##

    @staticmethod
    def gen_progress(iterable, title="Выполнение операции..."):
        """ Returns progress bar generator """
        return track(iterable, description=title)

    def message_callback(self, message, status='ok'):
        """ A complete message callback for FileController """
        color = 'green'
        if status == 'info': color = 'white'
        elif status == 'bad': color = 'red'
        elif status == 'nostatus': color = ''

        if type(message) == list and len(message) > 0 and type(message[0]) == UserAction:
            table = Table()
            table.add_column('Действие')
            table.add_column('Статус')
            for i, ua in enumerate(message):
                table.add_row(f"[dim]{i+1}[/dim] " + ua.descr(), "[green]OK" if ua.completed else "[red]ERROR")
            self.print(table)
            return

        if color:
            self.print(f"[{color}][not bold]{message}[/not bold][/{color}]")
        else:
            self.print(message)

    def parse_args(self):
        if '--test' in sys.argv:
            self.test()
            return

        if '--settings' in sys.argv:
            print(Settings.get_filepath())
            return
        
        if '--version' in sys.argv:
            import version
            print(Text(version.VERSION))
            return

        try:
            self.run_menu()
        except KeyboardInterrupt:
            pass

    def run_action_show_user_info(self) -> Optional[str]:
        emails = self.ask("Введите email").split(', ')
        if emails[0] == '!step1':
            emails = Settings()['cache.file_step1_emails']
            if emails:
                self.print(">>", Text(emails))
                emails = emails.split(', ')
            else:
                self.print("[red]Кэшированные email's отсутствуют.")
                input("Нажмите Enter чтобы продолжить...")
                return
        self.print("[dim]Введите путь к файлу, если хотите подгрузить данные из таблицы. Иначе введите «!»")
        filepath = self.ask_filepath(required=False)
        data_to_show = list()  # Содержит элементы вида (запрашиваемый email, List[UserInfo]|None)
        
        status = self.status("Получение информации о пользователе... ")
        status.start()
        learning = self.create_learning()
        driver = ExcelDriver()
        if filepath != "!":
            driver.load(filepath)

        # Загрузка данных пользователей
        try:
            for email in emails:
                try:
                    uinfo = learning.get_user_info(email)
                    if filepath != '!':
                        try:
                            for _uinfo in uinfo:
                                _uinfo.table = driver.get_user_data(email)
                        except ExcelUserNotFound:
                            pass
                    data_to_show.append((email, uinfo))
                except UserNotFound:
                    data_to_show.append((email, None))
        except NotAuthorized:
            status.stop()
            self.print("[red]Сессия устарела. Необходимо пройти аутентификацию ещё раз")
            del Settings()[self.AUTHCOOKIEID]
            return "rerender menu"
        status.stop()
        
        # Вывод загруженных данных
        data_to_show_len = len(data_to_show)
        for i, data in enumerate(data_to_show):
            title = None
            if data_to_show_len > 1:
                title = f"[cyan][not dim]Пользователь {data[0]} [bold]({i+1}[/bold]/[bold]{data_to_show_len})"
            if not data[1]:
                self.print(f"[magenta]Пользователь {data[0]} [bold]({i+1}[/bold]/[bold]{data_to_show_len})[/bold] не найден")
                val = input("Нажмите Enter чтобы продолжить...")
                if val == '!!': return
                continue
            for uinfo in data[1]:
                self.print()
                self.print(self.compose_user_info(uinfo, detailed=True, title=title))
                val = input("Нажмите Enter чтобы продолжить...")
                if val == '!!': return
        ##

    def run_action_perform_actions(self) -> Optional[str]:
        emails = self.ask("Введите email").split(', ')
        self.print("[dim]Введите путь к файлу, если хотите подгрузить данные из таблицы. Иначе введите «!»")
        filepath = self.ask_filepath(required=False)
        data_to_show = list()  # Содержит элементы вида (запрашиваемый email, List[UserInfo]|None)
        
        status = self.status("Получение информации о пользователе... ")
        status.start()
        abilities = ['learning']
        learning = self.create_learning()
        driver = ExcelDriver()
        if filepath != "!":
            driver.load(filepath)
            abilities.append('excel')
        

        # Загрузка данных пользователей
        try:
            for email in emails:
                try:
                    uinfo = learning.get_user_info(email)
                    if filepath != '!':
                        try:
                            for _uinfo in uinfo:
                                _uinfo.table = driver.get_user_data(email)
                        except ExcelUserNotFound:
                            pass
                    data_to_show.append((email, uinfo))
                except UserNotFound:
                    data_to_show.append((email, None))
        except NotAuthorized:
            status.stop()
            self.print("[red]Сессия устарела. Необходимо пройти аутентификацию ещё раз")
            del Settings()[self.AUTHCOOKIEID]
            return "rerender menu"
        status.stop()
        
        # Вывод загруженных данных
        data_to_show_len = len(data_to_show)
        for i, data in enumerate(data_to_show):
            title = None
            if data_to_show_len > 1:
                title = f"[cyan][not dim]Пользователь {data[0]} [bold]({i+1}[not bold]/[bold]{data_to_show_len})"
            if not data[1]:
                self.print(f"[magenta]Пользователь {email} не найден")
                val = input("Нажмите Enter чтобы продолжить...")
                if val == '!!': return
                continue
            for uinfo in data[1]:
                suggested = []
                if 'excel' in abilities:
                    with self.status("Выбор действий... "):
                        suggested = suggest_user_actions(uinfo, learning=learning)
                uactions = self.select_user_actions(uinfo, suggested, abilities, add_top_gap=True)
                try:
                    with self.status("Выполнение... "):
                        FileController.perform_user_actions(driver, learning, uinfo, uactions)
                    self.print("[green]Действия выполнены")
                except:
                    self.message_callback(traceback.format_exc(), status="info")
                    self.message_callback('Не удалось выполнить действия', 'bad')
                    self.message_callback(uactions, status='info')
        ##

    def run_menu(self):

        menu, menuitems, menuitems_len = self._create_menu()

        _is_first_run = True
        while True:
            if not _is_first_run: self.print("\n")
            self.print(menu)
            selection = IntPrompt.ask("[yellow]Выберите действие")

            # Выполнение действия
            action = None
            if selection > 0 and selection <= menuitems_len:
                action = menuitems[selection - 1].action

            if action == 'auth_check':
                self.print("")
                self.elearning_auth_check()
                menu, menuitems, menuitems_len = self._create_menu()  # update menu
            elif action == 'process_file_1':
                filepath = self.ask_filepath()
                
                confirm = Confirm.ask("[red]Данные в файле будут изменены или удалены. Продолжить?")
                if not confirm: 
                    _is_first_run = False
                    continue

                try:
                    FileController.step1(
                        filepath=filepath, 
                        progress_gen=self.gen_progress,
                        ask_user_actions=lambda *args, **kwargs: self.select_user_actions(*args, abilities=('learning', 'excel'), add_top_gap=True, **kwargs),
                        confirm_users_actions=self.confirm_users_actions,
                        message_callback=self.message_callback
                    )
                except RequestError as e:
                    self.print("\n[bold red]Ошибка запроса.[/bold red] Текст ошибки:\n")
                    self.print(e)
                    exit()
                except NotAuthorized:
                    self.print("\n[red]Сессия устарела. Необходимо пройти аутентификацию ещё раз")
                    del Settings()[self.AUTHCOOKIEID]
                    menu, menuitems, menuitems_len = self._create_menu()  # update menu
            elif action == 'process_file_2':
                filepath = self.ask_filepath()
                FileController.step2(filepath, self.message_callback)
            elif action == 'show_uinfo':
                msg = self.run_action_show_user_info()
                if msg == "rerender menu":
                    menu, menuitems, menuitems_len = self._create_menu()  # update menu
                ##
            elif action == 'perform_actions':
                msg = self.run_action_perform_actions()
                if msg == "rerender menu":
                    menu, menuitems, menuitems_len = self._create_menu()  # update menu
                ##
            elif action == 'find_password':
                mid = self.ask_int("Введите id пользователя")
                with self.status("Поиск пароля"):
                    learning = self.create_learning()
                    passw = learning.get_user_password(mid)
                if passw: self.print('[blue]' + passw)
                else: self.print('[magenta]Пароль не найден')
            elif action == 'logout':
                learning = self.create_learning()
                with self.status("Выход из системы... "):
                    learning.logout()
                    del Settings()[self.AUTHCOOKIEID]
                    menu, menuitems, menuitems_len = self._create_menu()  # update menu
            elif action == 'manage_labels':
                self.run_labels_menu()
            elif action == 'exit':
                exit()
            _is_first_run = False
            
    def run_labels_menu(self):
        menu, menuitems, menuitems_len = self._create_labels_menu()

        while True:
            self.print("\n")
            self.print(menu)
            selection = IntPrompt.ask("[yellow]Выберите действие")

            # Выполнение действия
            action = None
            if selection > 0 and selection <= menuitems_len:
                action = menuitems[selection - 1].action

            match action:
                case 'add':
                    self.print("\n[dim blue]Добавление экзамена[/dim blue]")
                    exam = self.ask_exam_params()
                    LabelController.add_exams([exam])
                case 'show_all':
                    exams = LabelController.get_exams()
                    self.print(self.compose_exams_table(exams))
                    input("Нажмите Enter чтобы продолжить...")
                case 'edit':
                    self.print("\n[dim blue]Редактирование экзамена[/dim blue]\n")
                    exams = LabelController.get_exams()
                    exam_id = -1 + self.ask_exam_number(exams)
                    exam = exams[exam_id]
                    new_exam = self.ask_exam_params(exam=exam)
                    LabelController.edit_exam(exam, new_exam)
                    self.print("[green]OK")
                case 'delete':
                    self.print("\n[dim blue]Удаление экзамена[/dim blue]\n")
                    exams = LabelController.get_exams()
                    exam_id = -1 + self.ask_exam_number(exams)
                    LabelController.delete_exam(exams[exam_id])
                    self.print("[green]OK")
                case 'test': 
                    self.print("\n[dim blue]Определение метки")
                    subject = self.ask("[cyan]Введите название предмета")
                    date = self.ask("[cyan]Введите дату прохождения экзамена[/cyan] [dim](ДД.ММ, «!» для текущей)")
                    label = ""
                    try:
                        if date == "!":
                            label = LabelController.get_label(subject)
                        else:
                            date_parts = date.split('.')
                            label = LabelController.get_label(subject, datetime.date(2000, int(date_parts[1]), int(date_parts[0])))
                        self.print("[magenta]>> Метка:", label)
                    except NoSuitableLabelFound:
                        self.print("[magenta]>> Метка не найдена")
                    input("Нажмите Enter чтобы продолжить...")
                case 'share':
                    self.print("\n[dim blue]Обменивайтесь метками с помощью строки обмена!")

                    share_filepath = os.path.join(os.getcwd(), 'share_string.txt')
                    share_bytes = LabelController.get_exams_share_bytes()
                    with open(share_filepath, mode='wb') as file: 
                        file.write(share_bytes)
                    self.print("[blue]Текущая строка обмена сохранена в файл по адресу:\n",
                        Text(share_filepath, style='yellow'))

                    selection = Prompt.ask(
                        "\n[cyan]Выберите действие[/cyan] [dim](! - выход, p - загрузить данные из файла обмена)",
                        choices=['!', 'p'], show_choices=False
                    )
                    match selection:
                        case 'c':
                            try:
                                copy_to_clipboard(share_string)
                                self.print("[green]Скопировано!")
                            except:
                                self.print("[red]Не удалось скопировать :(")
                            input("Нажмите Enter чтобы продолжить...")
                        case 'p':
                            filepath = self.ask_filepath(cache_suffix='share_labels')
                            self.print("[blue]Загрузка строки обмена... ")
                            share_bytes = None
                            with open(filepath, mode='rb') as file: 
                                share_bytes = file.read()
                            try:
                                LabelController.load_exams_from_share_bytes(share_bytes)
                            except:
                                self.print("[red]Возникла ошибка при разборе файла обмена")
                                raise
                            self.print("[green]Загрузка данных из файла обмена завершена.")

                case 'exit': return

    def select_user_actions(self, uinfo, suggested: List[UserAction]=[], 
            abilities: Iterable = (), add_top_gap=False) -> List[UserAction]:
        selection = None
        detailed_view = False
        is_first_run = True
        suggested_len = len(suggested)

        uactions_return = suggested
        if suggested_len: uactions_return.sort(key=lambda x: x.sort_key)

        uactions_return_add = lambda x: uactions_return.append(x) if x not in uactions_return else None
        error = ""

        def print_uactions():
            self.print(Panel.fit(
                "\n".join(
                    tuple(f"[dim]{i+1}[/dim] " + x.descr() for i,x in enumerate(uactions_return))
                ),
                title="[dim][italic]Выбранные действия"
            ))

        while True:
            uactions_return_len = len(uactions_return)

            # Вывод информации о пользователе
            if add_top_gap or is_first_run == False: self.print("\n")
            self.print(self.compose_user_info(uinfo, detailed_view))

            # Вывод выбранных действий
            if uactions_return_len > 0:
                print_uactions()

            # Вывод ошибок
            if error: self.print(error); error = ""
            
            # Выбор действия
            if is_first_run and suggested_len == 0:
                uactions_return.append(UserAction.SILENT_SKIP)
                uactions_return_len += 1
                print_uactions()
                self.print("[green]Предложенных действий нет[/green]")
            try:
                selection = self.ask_user_action(abilities, with_help=(selection == "?"))
            except KeyboardInterrupt:
                selection = None
            
            # Определение действия
            if (selection == '!'): 
                if uactions_return_len > 0:
                    if UserAction.DELETE in uactions_return: 
                        uactions_return = [UserAction.DELETE]
                    return uactions_return
                else:
                    error = "[red]Err: Необходимо выбрать действие[/red]"
            elif (selection == '!!'): exit()
            elif (selection == 'c'):
                comment = self.ask("Введите текст комментария")
                uactions_return_add(UserAction(UserActionType.SET_COMMENT, comment))
            elif (selection == 'd'):
                uactions_return.clear()
                uactions_return_add(UserAction.DELETE)
            elif (selection == 'dt'):
                uactions_return_add(UserAction.DELETE_FROM_TABLE)
            elif (selection == 'k'): 
                if UserAction.SKIP in uactions_return:
                    uactions_return.remove(UserAction.SKIP)
                uactions_return_add(UserAction.SILENT_SKIP)
            elif (selection == 'l'): uactions_return_add(UserAction.CHANGE_LOGIN)
            elif (selection == 'm'):
                label = self.ask("Введите название метки")
                uactions_return_add(UserAction(UserActionType.ADD_LABEL, label))
            elif (selection == 'md'):
                label = self.ask("Введите название метки")
                uactions_return_add(UserAction(UserActionType.REMOVE_LABEL, label))
            elif (selection == 'pe'):
                validator = None
                if UserActionType.CHANGE_PASSW_LOCAL in uactions_return:
                    idx = uactions_return.index(UserActionType.CHANGE_PASSW_LOCAL)
                    validator = lambda x: self.compare_passwords(x, uactions_return[idx].param)
                passw = self.ask("Введите пароль", validator=validator)
                if UserActionType.CHANGE_PASSW_EDU in uactions_return:
                    uactions_return.remove(UserActionType.CHANGE_PASSW_EDU)
                uactions_return_add(UserAction(UserActionType.CHANGE_PASSW_EDU, passw))
            elif (selection == 'pl'):
                validator = None
                if UserActionType.CHANGE_PASSW_EDU in uactions_return:
                    idx = uactions_return.index(UserActionType.CHANGE_PASSW_EDU)
                    validator = lambda x: self.compare_passwords(x, uactions_return[idx].param)
                passw = self.ask("Введите пароль", validator=validator)
                if UserActionType.CHANGE_PASSW_LOCAL in uactions_return:
                    uactions_return.remove(UserActionType.CHANGE_PASSW_LOCAL)
                uactions_return_add(UserAction(UserActionType.CHANGE_PASSW_LOCAL, passw))
            elif (selection == 'pg'):
                login = uinfo.table.login if uinfo.table and uinfo.table.login else ""
                passw = generate_password(login)
                changed = False
                if UserActionType.CHANGE_PASSW_LOCAL in uactions_return:
                    uactions_return.remove(UserActionType.CHANGE_PASSW_LOCAL)
                if UserActionType.CHANGE_PASSW_EDU in uactions_return:
                    uactions_return.remove(UserActionType.CHANGE_PASSW_EDU)
                if 'excel' in abilities:
                    uactions_return_add(UserAction(UserActionType.CHANGE_PASSW_LOCAL, passw))
                    changed = True
                if 'learning' in abilities:
                    uactions_return_add(UserAction(UserActionType.CHANGE_PASSW_EDU, passw))
                    changed = True
                if changed:
                    self.print("[green]Действия по смене пароля добавлены.")
                    input("Нажмите Enter чтобы продолжить...")
            elif (selection == 'r'): uactions_return.clear()
            elif (selection == 'ro'): 
                idx = -1 + self.ask_int("Введите номер действия", 1, len(uactions_return))
                uactions_return.pop(idx)
            elif (selection == 's'): 
                if UserAction.SILENT_SKIP in uactions_return:
                    uactions_return.remove(UserAction.SILENT_SKIP)
                uactions_return_add(UserAction.SKIP)
            elif (selection == 'sr'): 
                uactions_return_add(UserAction.MARK_REGISTERED)
            elif (selection == 'x'): detailed_view = not detailed_view
            uactions_return.sort(key=lambda x: x.sort_key)
            is_first_run = False

        # сюда ход выполнения доходить не должен
        if UserAction.DELETE in uactions_return: uactions_return = [UserAction.DELETE]
        return uactions_return

    def test(self):
        from utils import is_red_color
        print(is_red_color('ff0000'))
        print(is_red_color('f40d32'))
        print(is_red_color('e36b09'))
        print(is_red_color('dbd40d'))
        