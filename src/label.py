import base64
import datetime
from typing import NamedTuple, Optional, List, Tuple, Iterable
import json

import custom_json
from settings import Settings
from datatypes import Exam, Label

class LabelControllerError(Exception):
    pass

class DataIsNotEnoughError(LabelControllerError):
    pass

class NoSuitableLabelFound(LabelControllerError):
    pass


class LabelController:

    _all_exams: List[Exam] = list()
    _pr = 'LabelController'

    @classmethod
    def add_exams(cls, exams: Iterable[Exam]) -> None:
        """ Add exams to list """
        cls.load_exams()
        for exam in exams:
            if exam in cls._all_exams: continue
            cls._all_exams.append(exam)
        Settings()[f"{cls._pr}.exams"] = cls._all_exams

    @classmethod
    def delete_exam(cls, exam: Exam) -> None:
        cls.load_exams()
        cls._all_exams.remove(exam)
        Settings()[f"{cls._pr}.exams"] = cls._all_exams

    @classmethod
    def edit_exam(cls, old: Exam, new: Exam) -> bool:
        """ Works with loaded exams """
        cls.load_exams()
        idx = cls._all_exams.index(old)
        cls._all_exams[idx] = new
        Settings()[f"{cls._pr}.exams"] = cls._all_exams
        return True

    @classmethod
    def load_exams(cls) -> None:
        _exams = list()

        exams = Settings()[f"{cls._pr}.exams"]
        if exams == cls._all_exams: return
        if not exams: exams = list()
        
        for exam in exams:
            dates = [datetime.date.fromisoformat(x) for x in exam[2]]
            _exams.append(Exam(exam[0], exam[1], dates))
        cls._all_exams = _exams

    @classmethod
    def load_exams_from_share_bytes(cls, share: bytes) -> None:
        s = base64.b64decode(share).decode()
        Settings()[f"{cls._pr}.exams"] = json.loads(s)
        cls.load_exams()

    @classmethod
    def get_exams(cls) -> Tuple[Exam]:
        cls.load_exams()
        return tuple(cls._all_exams)

    @classmethod
    def get_exams_share_bytes(cls) -> bytes:
        cls.load_exams()
        json_str = json.dumps(cls._all_exams, cls=custom_json.JSONEncoder)
        return base64.b64encode(bytes(json_str, 'ascii'))
    
    @classmethod
    def get_label(cls, exam_subject: str, selected_date: datetime.date | None = None) -> str:
        """Определеяет метку для экзамена (с ближайшим блоком)

        Экзамен переносится на следующий блок, если до экзамена один день
        или экзамен проводится во время начала блока

        :param exam_subject: Название экзамена
        :type exam_subject: str
        :param selected_date: Выбранная дата сдачи экзамена
        :type selected_date: datetime.date
        :raises NoSuitableLabelFound: Если метка не определена
        :returns: Возвращает метку для экзамена (с блоком)
        :rtype: str
        """

        cls.load_exams()

        ##
        current_date = datetime.date.today()
        current_year = current_date.year
        current_date = current_date.replace(year=2000)
        day = datetime.timedelta(days=1)
        selected_date = selected_date.replace(year=2000) if selected_date else None
        ##

        for exam in cls._all_exams:
            if exam.subject.lower() != exam_subject.lower():
                continue
            block = 0
            for date_i, date in enumerate(exam.dates):
                if selected_date == date:
                    return f"{exam.tag}{current_year}{date_i+1}"
                elif selected_date == None and current_date < date - day:
                    return f"{exam.tag}{current_year}{date_i+1}"

        raise NoSuitableLabelFound(f"exam_subject={exam_subject}, selected_date={selected_date}")

    @classmethod
    def get_label_primitive(cls, exam_subject: str) -> Label:
        """ Временный аналог get_label """
        exam_labels = {
            'биология': 'Био20252',

            'иностранный язык': 'Иняз20252',
            
            'информатика и икт': 'Инф20252',
            'информатика, алгоритмизация и программирование': 'ИнфП20252',
            
            'история': 'Ист20252',
            'история государства и общества': 'ИстП20252',
            
            'математика': 'Мат20252',
            'математика (английский)': 'МатА20252',
            'математика (англ)': 'МатА20252',
            'математика в технике и технологиях': 'МатТ20252',
            'математика в управлении': 'МатЭ20252',
            
            'обществознание': 'Общ20252',
            'обществознание (основы общественных наук)': 'ОбщП20252',
            
            'русский язык для иностранцев': 'РусИн20252',
            'русский язык для иностранных граждан': 'РусИн20252',
            'русский язык': 'Рус20252',
            
            'физика': 'Физ20252',
            'общая физика': 'ФизП20252',
            
            'химия': 'Хим20252',
            'инженерная химия': 'ХимП20252',
        }

        label = exam_labels.get(str(exam_subject).lower())
        if not label:
            raise DataIsNotEnoughError
        
        tag = label[:-1]
        block = int(label[-1:])
        return Label(exam_subject, tag, block)

    @classmethod
    def test(cls):
        from collections import namedtuple
        from typing import NamedTuple
        selected_date = datetime.date(2025, 8, 18)
        # Exam = namedtuple('Exam', ('subject', 'tag', 'dates'))
        Exam = NamedTuple('Exam', (('subject', str), ('tag', str), ('dates', List[datetime.date])))
        exams = [
            Exam('химия', 'Хим', [datetime.date(2000, 6, 18), datetime.date(2000, 7, 18), datetime.date(2000, 8, 18)]),
            Exam('физика', 'Физ', [datetime.date(2000, 6, 20), datetime.date(2000, 7, 20), datetime.date(2000, 8, 20)]),
        ]
        inp = 'химия'
        current_date = datetime.date.today()
        current_year = current_date.year
        current_date = current_date.replace(year=2000)
        day = datetime.timedelta(days=1)
        selected_date = selected_date.replace(year=2000) if selected_date else None
        output = []

        for e in exams:
            if e.subject.lower() != inp.lower():
                continue
            block = 0
            for date_i, date in enumerate(e.dates):
                if selected_date == date:
                    output.append(f"{e.tag}{current_year}{date_i+1}")
                elif selected_date == None and current_date < date - day:
                    output.append(f"{e.tag}{current_year}{date_i+1}")
        print(output)

    @classmethod
    def test2(cls):
        print(cls.get_label('химия'))

    


        
    