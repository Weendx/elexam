import datetime
from dataclasses import dataclass
from enum import Enum
from collections import namedtuple
from typing import List, NamedTuple, Optional, Tuple

AuthCookies = namedtuple('AuthCookies', ('PHPSESSID', 'hmkey'))

@dataclass
class Course:
    cid: int  # course id
    title: str
    starts: Optional[datetime.datetime] = None
    ends: Optional[datetime.datetime] = None
    teachers: Optional[List[str]] = None

EmailNLogin = namedtuple('EmailNLogin', ('email', 'login'))

Exam = NamedTuple('Exam', (('subject', str), ('tag', str), ('dates', List[datetime.date])))

MenuItem = namedtuple('MenuItem', ['title', 'action'])

class Label(NamedTuple):
    subject: str  # предмет
    tag: str  # метка
    block: Optional[int] = None

    def __str__(self):
        label = self.tag
        if self.block: label += str(self.block)
        return label

class UserActionType(Enum):
    SKIP = 'skip'
    DELETE = 'delete'
    DELETE_FROM_TABLE = 'delete_from_table'
    ADD_LABEL = 'add_label'
    REMOVE_LABEL = 'remove_label'
    CHANGE_LOGIN = 'change_login'
    CHANGE_PASSW_LOCAL = 'change_password_local'
    CHANGE_PASSW_EDU = 'change_password_edu'
    MARK_REGISTERED = 'mark_registered'
    SILENT_SKIP = 'silent_skip'
    SET_COMMENT = 'set_comment'

class UserActionMeta(type):
    def __getattr__(cls, name):
        if name in UserActionType.__members__:  # Проверяем, есть ли такой элемент в UserActionType
            type_member = UserActionType[name]
            return cls(type_member) 
        raise AttributeError(f"'{cls.__name__}' object has no attribute '{name}'")

class UserAction(metaclass=UserActionMeta):

    def __init__(self, action: UserActionType, param: str = ""):
        if not isinstance(action, UserActionType):
            raise AttributeError(f"action attribute must be UserActionType instance")
        
        param_required = (
            UserActionType.ADD_LABEL, UserActionType.REMOVE_LABEL,
            UserActionType.SET_COMMENT,
            UserActionType.CHANGE_PASSW_EDU, UserActionType.CHANGE_PASSW_LOCAL
        )

        if not param and action in param_required:
            raise AttributeError(f"param should be passed when action is {action.name}")

        self.action = action
        self.param = param
        self.weight = 50
        self.completed = False
        self.requires = ()

        match action:
            case UserActionType.DELETE: self.weight = 100
            case UserActionType.SET_COMMENT: self.weight = 90
            case UserActionType.MARK_REGISTERED: self.weight = 85
            case UserActionType.SKIP: self.weight = 85
            case UserActionType.CHANGE_LOGIN: self.weight = 10
            case UserActionType.CHANGE_PASSW_EDU: self.weight = 11
            case UserActionType.CHANGE_PASSW_LOCAL: self.weight = 20

        match action:
            case UserActionType.SKIP: self.requires = ('excel')
            case UserActionType.DELETE: self.requires = ('learning')
            case UserActionType.DELETE_FROM_TABLE: self.requires = ('excel')
            case UserActionType.ADD_LABEL: self.requires = ('learning', 'excel')
            case UserActionType.REMOVE_LABEL: self.requires = ('learning')
            case UserActionType.CHANGE_LOGIN: self.requires = ('learning', 'excel')
            case UserActionType.CHANGE_PASSW_LOCAL: self.requires = ('excel')
            case UserActionType.CHANGE_PASSW_EDU: self.requires = ('learning')
            case UserActionType.MARK_REGISTERED: self.requires = ('excel')
            case UserActionType.SILENT_SKIP: self.requires = ()
            case UserActionType.SET_COMMENT: self.requires = ('excel')

    def __eq__(self, other):
        if isinstance(other, UserActionType):
            return self.action == other
        if isinstance(other, self.__class__):
            ignore_params = (
                UserActionType.SET_COMMENT, 
                UserActionType.CHANGE_PASSW_EDU,
                UserActionType.CHANGE_PASSW_LOCAL
            )
            if self.action in ignore_params:
                return self.action == other.action
            return self.action == other.action and self.param == other.param
        return False
    
    def __hash__(self):
        return hash((self.action, self.param))

    def __repr__(self):
        param = f":{self.param!r}" if self.param else ''
        completed = " COMPLETED" if self.completed else ''
        return f"<UserAction {self.action.name}{param}{completed}>" 

    def __str__(self):
        return self.descr()

    @property
    def sort_key(self):
        return self.weight

    def descr(self):
        if self.action == UserActionType.SKIP: return "Пропустить с пометкой в таблице"
        elif self.action == UserActionType.DELETE: return "Удалить пользователя"
        elif self.action == UserActionType.DELETE_FROM_TABLE: return "Удалить пользователя из таблицы"
        elif self.action == UserActionType.ADD_LABEL: return f"Добавить метку <[dodger_blue3]{self.param}[/dodger_blue3]>"
        elif self.action == UserActionType.REMOVE_LABEL: return f"Убрать метку <[dodger_blue3]{self.param}[/dodger_blue3]>"
        elif self.action == UserActionType.CHANGE_LOGIN: return "Записать в таблицу логин из eLearning"
        elif self.action == UserActionType.CHANGE_PASSW_EDU: return f"Установить пароль <[dodger_blue3]{self.param}[/dodger_blue3]> в eLearning"
        elif self.action == UserActionType.CHANGE_PASSW_LOCAL: return f"Установить пароль <[dodger_blue3]{self.param}[/dodger_blue3]> в таблице"
        elif self.action == UserActionType.MARK_REGISTERED: return "Отметить в таблице, что зарегистрирован"
        elif self.action == UserActionType.SILENT_SKIP: return "Пропустить без пометки в таблице"
        elif self.action == UserActionType.SET_COMMENT: return f"Установить комментарий <[dodger_blue3]{self.param}[/dodger_blue3]>"
        return None

class TableSubject(NamedTuple):
    name: str
    date: datetime.datetime | None = None

class UserTableData(NamedTuple):
    email: str
    login: str
    fio: Optional[str] = None
    subjects: List[TableSubject] = []

@dataclass
class UserInfo:
    mid: int
    login: str
    email: str
    fio: str
    tags: Optional[Tuple[str]] = None
    table: Optional[UserTableData] = None
    registered: Optional[datetime.datetime] = None
    last_login: Optional[datetime.datetime] = None
    courses: Optional[Tuple[Course]] = None
    source: Optional[str] = 'elexam'

