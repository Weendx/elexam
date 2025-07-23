
# selectolax for html parsing
from datetime import datetime
from collections import namedtuple
import html
import requests
import json
from selectolax.parser import HTMLParser
from typing import Tuple, List, Union
import urllib
import time

from datatypes import AuthCookies, UserInfo, Course
from utils import convert_date_string

from rich import print


class InvalidLoginPair(Exception):
    pass

class NotAuthorized(Exception):
    pass

class RequestError(Exception):
    pass

class UserNotFound(Exception):
    pass

class SomethingWrong(Exception):
    pass

class DataAgreementNotAccepted(Exception):
    pass


class LearningDriver:

    website = "https://edu.nntu.ru"

    _auth_check_completed = False
    _current_role = None

    def __init__(self, auth_cookies: Union[AuthCookies, None] = None):
        self._session = requests.Session()
        self._session.headers.update({
            'IS_AJAX_REQUEST': 'TRUE',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0',
            # 'Content-Type': 'multipart/form-data'
        })
        
        if auth_cookies: 
            self._session.cookies.set(
                name='PHPSESSID',
                value=auth_cookies.PHPSESSID,
                domain='edu.nntu.ru',
                path='/',
                secure=True
            )
            self._session.cookies.set(
                name='hmkey',
                value=auth_cookies.hmkey,
                domain='edu.nntu.ru',
                path='/',
                secure=True
            )
            
        self._session.cookies.set(
            name='hmlang',
            value='rus',
            domain='edu.nntu.ru',
            path='/',
            secure=True
        )

    def _auth_check(self):
        if not self.auth_check():
            raise NotAuthorized

    def add_tag(self, user_id, tag):
        self._auth_check()
        params = {
            'postMassIds_grid': user_id, 
            'massActionsAll_grid': user_id,
            'tags[]': tag
        }
        
        # response is dumb
        self.request('/user/list/assign-tag', params, 'post')

    def auth(self, login, password):
        """ Returns hmkey cookie value """
        # remove old cookies
        self._session.cookies.clear()

        params = {
            'ref': '/index/login', 'start_login': 1,
            'login': login, 'password': password,
            'remember': 1, 'form_redirectUrl': ''
        }

        resp = self.request('/index/authorization', params=params, method='post')
        if resp.get('code') == 1:
            return AuthCookies(
                self._session.cookies.get('PHPSESSID'), 
                self._session.cookies.get('hmkey')
            )
        else:
            raise InvalidLoginPair()

    def auth_check(self) -> bool:
        if not self._session.cookies.get('PHPSESSID'): return False
        if self._auth_check_completed: return True
        resp = self.request('/')
        if resp == True:
            self._auth_check_completed = True
            return True
        return False

    def delete(self, user_id) -> bool:
        self._auth_check()

        params = {'postMassIds_grid': user_id, 'massActionsAll_grid': user_id}
        self.request('/user/list/delete-by', params, 'post')
        # успешно!
        is_ajax_request_value = self._session.headers.pop('IS_AJAX_REQUEST', None)

        page = self.request('/user/list/delete-by', params, 'post', format_='html')
        
        if is_ajax_request_value:
            self._session.headers.update({'IS_AJAX_REQUEST': is_ajax_request_value})
        
        msg = self.get_notification(page)

        if "успешно" in msg:
            return True
        else:
            raise SomethingWrong(f"Something wrong on user {user_id} deleting: " + repr(msg))

    def get_current_info(self):
        self._auth_check()

        # caching
        if hasattr(self, '_get_current_info_last') and hasattr(self, '_get_current_info_return') \
             and self._get_current_info_last and self._get_current_info_return \
             and time.time() - self._get_current_info_last < 2:
            return self._get_current_info_return
            
        resp = self.request('/user/ajax/current-user-data')
        self._get_current_info_last = time.time()
        self._get_current_info_return = resp
        return resp
    
    def get_current_fio(self) -> str:
        resp = self.get_current_info()
        name = list()
        if resp.get('LastName'): name.append(resp['LastName'])
        if resp.get('FirstName'): name.append(resp['FirstName'])
        if resp.get('Patronymic'): name.append(resp['Patronymic'])
        return " ".join(name)

    def get_current_role(self) -> str:
        resp = self.get_current_info()
        return resp['role'] if resp.get('role') else ''

    def get_current_roles(self) -> str:
        resp = self.get_current_info()
        return resp['roles'] if resp.get('roles') else dict()
    
    def get_log_message(self, log_id) -> str | None:
        """ Returns html-like str if message found, else None """
        if not log_id: return None
        resp = self.request(f"/notice/log/one/log_id/{log_id}")
        if not resp: return None
        fields = resp.get('fields')
        if not fields: return None
        for field in fields:
            if field.get('key') == 'Сообщение' or field.get('key') == 'Message':
                return field.get('value')
        ##
        return None

    @classmethod
    def get_notification(cls, html_: str) -> None | str | dict | list:
        """ Returns text from notification box

            Args:
                html_(str): page html
            Returns:
                None: if there is no messages
                str: if one message
                dict: ({'message':'...','type':3}) if message is specific
                list: if more than one message
        """
        parser = HTMLParser(html.unescape(html_))
        hm_notif_attrs = parser.css_first('hm-notifications').attributes
        
        notifications = hm_notif_attrs.get(':notifications')
        if not notifications: return None
        notifications = json.loads(notifications)
        
        if not notifications:
            return None
        elif type(notifications) == list and len(notifications) == 1:
            return notifications[0]
        else:
            return notifications

    def get_user_courses(self, user_id) -> Tuple[Course]:
        """ Returns user courses """
        
        self._auth_check()

        params = {
            'gridmod': 'ajax', 'grid': 'grid',
            'page': 1, 'perPage': 30, 'ordergrid': 'subjectId_ASC',
            'personId': user_id
        }
        resp = self.request(f"/report/index/index/report_id/29", params=params, method='post')
        data = resp.get('data')
        courses = dict()

        for row in data:
            person_id = row.get('personId')
            if int(person_id) != int(user_id): continue
            course_id = row.get('subjectId')
            if not course_id: continue
            course_id = int(course_id)
            
            course = courses.get(course_id)
            if not course:
                course = Course(course_id, title = html.unescape(row['subjectTitle'].replace('&amp;', '&')))
                courses[course_id] = course
            
            if not course.starts: 
                course.starts = convert_date_string(row.get('subjectBegin'))
            if not course.ends:
                course.ends = convert_date_string(row.get('subjectEnd'))
                
            teacher = row.get('teacherFio')
            if teacher:
                if course.teachers == None: course.teachers = list()
                course.teachers.append(teacher)
        ####
        return tuple(courses.values())
                
    def get_user_info(self, email, load_courses=True) -> List[UserInfo]:
        """ Returns user info """

        self._auth_check()
        
        # Переключаемся на администратора
        if not self.switch_role('admin'):
            raise SomethingWrong('Got error while switching to admin role')

        params = {
            'gridmod': 'ajax', 'grid': 'grid',
            'perPage': 30, 'page': 1, 'ordergrid': 'fio_ASC',
            'email': email
        }
        resp = self.request('/user/list', params=params, method='post')
        
        data = resp.get('data')

        if not data:
            if type(data) == list:
                raise UserNotFound
            raise UserNotFound(repr(resp))
        
        uinfo_return = list()

        for row in data:
            if str(row['email']).lower() != str(email).lower(): continue
            uinfo = UserInfo(
                mid = int(row['MID']),
                login = row['login'],
                email = row['email'],
                fio = HTMLParser(html.unescape(row['fio'])).css_first('a').text(),
            )

            if row.get('Registered'):
                uinfo.registered = convert_date_string(row['Registered'])
            if row.get('last_login_date'):
                uinfo.last_login = convert_date_string(row['last_login_date'])
            if row.get('tags'):
                tree = HTMLParser(html.unescape(row['tags']))
                tags = [x for x in map(lambda x: x.text(), tree.css('p'))]
                if (len(tags) > 1): tags = tags[1:]
                uinfo.tags = tuple(tags)
            if row.get('source'):
                uinfo.source = row['source']

            if load_courses:
                uinfo.courses = self.get_user_courses(uinfo.mid)
            
            uinfo_return.append(uinfo)

        if not uinfo_return:
            raise UserNotFound
        return uinfo_return

    def get_user_password(self, user_id) -> str | None:
        """ Try to get user password from eLearning 
            Returns:
                Password, if found, else None
        """
        self._auth_check()
        
        params = {
            'gridmod': 'ajax', 'receiver_id': user_id,
            'cluster': 'general',  # Бизнес-процесс == Общего назначения
            'grid': 'grid', 'perPage': 30, 'page': 1,
            'ordergrid': 'send_date_DESC'
        }
        messages_general = self.request('/notice/log', params, 'post')
        general_data = messages_general.get('data')
        if not general_data: return None
        for general_data_row in general_data:
            receiver_id = general_data_row.get('receiver_id')
            if int(receiver_id) != int(user_id): continue
            theme = general_data_row.get('theme')
            if not 'Изменение пароля' in theme and not 'Вы зарегистрированы' in theme:
                continue
            # нужно получить самое последнее письмо с паролем
            # письма уже отсортированы по последнему в запросе
            # поэтому, получаем письмо и сразу возвращаем результат
            message = self.get_log_message(general_data_row.get('log_id'))
            if not message: return None  # Мб ошибка, отправляем на ручную проверку
            message_tree = HTMLParser(html.unescape(message))

            if 'Вы зарегистрированы' in theme:
                try:
                    # точка в конце в пароль не входит
                    password = message_tree.css('li')[1].text()  # ex: 'пароль - 1JXM.'
                    password = password[9:]  # ex: '1JXM.'
                    password = password[:-1]  # ex: '1JXM'
                    return password
                except:
                    return None
            elif 'Изменение пароля' in theme:
                try:
                    p_list = message_tree.css('p')
                    for p in p_list:
                        if 'Новый пароль: ' in p.text():  # ex: 'Новый пароль: el_6rBOw'
                            return p.text()[14:]  # ex: '1JXM'
                except:
                    return None
        return None

    def is_user_exists(self, email) -> bool:
        self._auth_check()

        # Переключаемся на администратора
        self.switch_role('admin')

        params = {
            'gridmod': 'ajax', 'grid': 'grid',
            'perPage': 30, 'page': 1, 'ordergrid': 'fio_ASC',
            'email': email
        }
        resp = self.request('/user/list', params=params, method='post')
        data = resp.get('data')
        if not data: return False
        for row in data:
            if str(row['email']).lower() == str(email).lower(): return True
        return False

    def logout(self):
        self.request('/logout')

    def remove_tag(self, user_id, tag):
        self._auth_check()
        params = {
            'postMassIds_grid': user_id, 
            'massActionsAll_grid': user_id,
            'tagsUnassign[]': tag
        }
        
        # response is dumb
        self.request('/user/list/unassign-tag', params, 'post')

    def request(self, endpoint, params={}, method='get', headers=None, format_="json"):
        try:
            if method == 'get':
                query = '?' + urllib.parse.urlencode(params) if not '?' in endpoint else ''
                resp = self._session.get(self.website + endpoint + query, headers=headers)
                # print('GET', resp.url)
                if format_ == "json":
                    resp = resp.json()
                else:
                    resp = resp.text
            elif method == 'post':
                resp = self._session.post(self.website + endpoint, data=params, headers=headers)
                # print('POST', resp.url)
                if format_ == "json":
                    resp = resp.json()
                else:
                    resp = resp.text
            else:
                raise AttributeError('Only get or post methods allowed')
        except requests.exceptions.ConnectionError as e:
            raise RequestError(str(e))

        if format_ == "json" and type(resp) == dict:
            if endpoint not in ['/logout', '/', '/index/authorization'] \
                    and resp.get('designOptions'):
                raise NotAuthorized
            elif resp.get('error'):
                raise RequestError
            elif resp.get('dataAgreement'):
                raise DataAgreementNotAccepted

        return resp

    def set_password(self, user_id, password) -> bool:
        self._auth_check()
        if not self.switch_role('admin'):
            raise SomethingWrong
        
        params = {
            'postMassIds_grid': user_id, 'massActionsAll_grid': user_id,
            'pass': password
        }
        
        # Пароль успешно назначен!
        is_ajax_request_value = self._session.headers.pop('IS_AJAX_REQUEST', None)

        page = self.request('/user/list/set-password', params, 'post', format_='html')
        
        if is_ajax_request_value:
            self._session.headers.update({'IS_AJAX_REQUEST': is_ajax_request_value})
        
        msg = self.get_notification(page)

        if "успешно" in msg:
            return True
        else:
            raise SomethingWrong("Something wrong on password change: " + repr(msg))

    def switch_role(self, role: str) -> bool:
        self._auth_check()
        if self._current_role == role: return True

        check = self.get_current_role()
        if check == role:
            self._current_role = role
            return True

        resp = self.request(f"/index/switch/role/{role}")
        if resp == True:
            check = self.get_current_role()
            if check != role:
                return False
            self._current_role = role
            return True
        return False

    