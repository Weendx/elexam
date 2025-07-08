from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet
import os
import re
import errno
from typing import List, Tuple, NamedTuple, Optional
from copy import copy

from datatypes import EmailNLogin, UserTableData, TableSubject
from utils import convert_date_string


class NotLoadedException(Exception):
    pass

class ColumnNotFoundException(Exception):
    pass

class UserNotFoundException(Exception):
    pass

class SheetNotFoundException(Exception):
    pass

class ExcelDriver:

    _xlsx = None
    _filepath = None

    def change_columns(self, email: str, changes: List[Tuple[str, str]]):
        """ Change values in columns
            changes: [(column_name, value), ...] 
        """

        self.check_loaded()

        for ws in self._xlsx:
            email_col = self.get_column_by_name(ws, 'email')
            column_numbers = dict()

            for change in changes:
                try:
                    column_numbers[change[0]] = self.get_column_by_name(ws, change[0])
                except ColumnNotFoundException:
                    continue

            for row in ws.iter_rows():
                email_cell = row[email_col - 1]            
                if email_cell.value != email: continue
                for change in changes:
                    col = column_numbers.get(change[0])
                    if not col: continue
                    cell = row[col - 1]
                    cell.value = change[1]

    def change_login_password(self, email: str, login: str, password: str):
        self.check_loaded()
        for sheetname in self._xlsx.sheetnames:
            ws = self._xlsx[sheetname]

            try:
                email_col = self.get_column_by_name(ws, 'email')
                login_col = self.get_column_by_name(ws, 'логин')
                passw_col = self.get_column_by_name(ws, 'пароль')
            except ColumnNotFoundException:
                continue

            max_col = max((email_col, login_col, passw_col))
            for row in ws.iter_rows(max_col=max_col):
                email_cell = row[email_col - 1]
                login_cell = row[login_col - 1]
                passw_cell = row[passw_col - 1]                
                if email_cell.value != email: continue
                login_cell.value = login
                passw_cell.value = password

    def check_loaded(self):
        if self._xlsx == None or self._filepath == None:
            raise NotLoadedException()

    def clone_sheet(self, source_worksheet) -> Worksheet:
        """ Returns worksheet with cloned data (unique) """
        return self._xlsx.copy_worksheet(source_worksheet)

    def clone_sheet_unique(self, ws_copy, ws_paste, unique_column_name) -> None:
        """ Copy data from ws_copy to ws_paste (unique) """
        unique_column = self.get_column_by_name(ws_copy, unique_column_name)
        email_column = self.get_column_by_name(ws_copy, 'email')
        password_column = self.get_column_by_name(ws_copy, 'пароль')
        password_formula = self.get_password_formula()

        formula_reexp = re.compile(r'=\(RIGHT\(..+(;|,)5\)\+23000\)\*15')

        added = list()
        last_inserted_row = 1
        for i, row in enumerate(ws_copy.iter_rows()):
            if ws_copy.cell(row=i+1, column=email_column).value == '<deleted>': continue
            unique_cell = ws_copy.cell(row=i+1, column=unique_column)
            if unique_cell.value in added: continue
            for cell in row:
                new_cell = ws_paste.cell(row=last_inserted_row, column=cell.column,
                        value=cell.value)
                if cell.column == password_column and i != 0 and formula_reexp.match(cell.value):
                        new_cell.value = password_formula.replace('%i', str(last_inserted_row))
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
            ##
            last_inserted_row += 1
            added.append(unique_cell.value)

    def create_sheet(self, *args, **kwargs):
        self.check_loaded()
        return self._xlsx.create_sheet(*args, **kwargs)

    def delete_user_from_workbook(self, email: str) -> None:
        self.check_loaded()
        for worksheet in self._xlsx:
            email_column = self.get_column_by_name(worksheet, 'email')
            for row in worksheet.iter_cols(min_col=email_column, max_col=email_column):
                for cell in row:
                    if cell.value != email:
                        continue
                    # worksheet.delete_rows(cell.row)
                    # worksheet.move_range(f'A{cell.row}:A{cell.row}', rows=-1, translate=True)
                    for cell_range in worksheet.iter_rows(min_row=cell.row, max_row=cell.row, max_col=9):
                        for cell_from_range in cell_range:
                            cell_from_range.value = None
                    cell.value = "<deleted>"
            
    def get_all_users_data(self, first_row_is_header=True) -> Tuple[UserTableData]:
        self.check_loaded()
        userdata = dict()
        ws = self.get_first_worksheet()

        email_column = self.get_column_by_name(ws, 'email')
        login_column = self.get_column_by_name(ws, 'логин')
        subject_column = self.get_column_by_name(ws, 'предмет')
        sel_date_column = self.get_column_by_name(ws, 'выбранная дата')
        fio_columns = self.get_columns_with_fio(ws)
        
        start_row = 2 if first_row_is_header else 1
        for column in ws.iter_cols(min_col=email_column, max_col=email_column, min_row=start_row):
            for cell in column:
                email = ws.cell(row=cell.row, column=email_column).value
                utd = userdata.get(email)
                if not utd:
                    fio = ""
                    for fio_subcolumn in fio_columns: 
                        value = ws.cell(row=cell.row, column=fio_subcolumn).value
                        if not value or not value.strip(): continue
                        if fio: value = ' ' + value
                        fio += value
                    if not fio: fio = None

                    userdata[email] = UserTableData(
                        email=email,
                        login=ws.cell(row=cell.row, column=login_column).value,
                        fio=fio,
                        subjects=[],
                    )
                    utd = userdata[email]
                subject_name = ws.cell(row=cell.row, column=subject_column).value
                subject_date = ws.cell(row=cell.row, column=sel_date_column).value
                subject_date = convert_date_string(str(subject_date)) if subject_date else None
                utd.subjects.append(TableSubject(subject_name, subject_date))

        return tuple(userdata.values())

    @classmethod
    def get_column_by_name(cls, worksheet, column_name) -> int:
        for row in worksheet.iter_rows(max_row=1):
            for cell in row:
                if not cell.value: continue
                if cell.value.lower().strip() == column_name.lower().strip():
                    return cell.column
        raise ColumnNotFoundException()
        
    @classmethod
    def get_columns_with_fio(cls, worksheet) -> Tuple[int]:
        """ Returns a tuple with fio columns
            Example: (2,3,4)
        """
        fio_last_name = cls.get_column_by_name(worksheet, 'фио')
        return (fio_last_name, fio_last_name+1, fio_last_name+2)

    def get_emails(self, worksheet, first_row_is_header=True):
        email_column = self.get_column_by_name(worksheet, 'email')
        start_row = 2 if first_row_is_header else 1
        email_generator = worksheet.iter_cols(min_row=start_row, min_col=email_column, 
                                                max_col=email_column, values_only=True)
        return set(next(email_generator))
    
    def get_emails_n_logins(self, worksheet, first_row_is_header=True):
        email_column = self.get_column_by_name(worksheet, 'email')
        login_column = self.get_column_by_name(worksheet, 'логин')
        max_col = email_column if email_column > login_column else login_column
        start_row = 2 if first_row_is_header else 1
        generator = worksheet.iter_rows(min_row=start_row, max_col=max_col, values_only=True)
        output = set()
        for row in generator:
            output.add(EmailNLogin(email=row[email_column - 1], login=row[login_column - 1]))
        return output
    
    def get_first_worksheet(self):
        self.check_loaded()
        ws = self._xlsx.active
        
        firstlist_names = ['лист 1', 'лист1', 'общий', 'sheet 1', 'sheet1']
        if ws.title.lower() in firstlist_names:
            return ws

        for sheetname in self._xlsx.sheetnames:
            if sheetname.lower() in firstlist_names:
                return self._xlsx[sheetname]
        raise SheetNotFoundException

    @staticmethod
    def get_password_formula() -> str:
        """ Returns a formula for password generation 
            
            Note: 
                use , as argument separator instead of ; in formula
                  %i for row number
            Example:
                =(RIGHT(H%i,5)+23000)*15
        """
        return "=(RIGHT(H%i,5)+23000)*15"

    def get_user_data(self, email, first_row_is_header=True) -> UserTableData:
        self.check_loaded()
        _email = email
        userdata = None
        ws = self.get_first_worksheet()

        email_column = self.get_column_by_name(ws, 'email')
        login_column = self.get_column_by_name(ws, 'логин')
        subject_column = self.get_column_by_name(ws, 'предмет')
        sel_date_column = self.get_column_by_name(ws, 'выбранная дата')
        fio_columns = self.get_columns_with_fio(ws)
        
        start_row = 2 if first_row_is_header else 1
        for column in ws.iter_cols(min_col=email_column, max_col=email_column, min_row=start_row):
            for cell in column:
                email = ws.cell(row=cell.row, column=email_column).value
                if email != _email: continue
                if not userdata:
                    fio = ""
                    for fio_subcolumn in fio_columns: 
                        value = ws.cell(row=cell.row, column=fio_subcolumn).value
                        if not value or not value.strip(): continue
                        if fio: value = ' ' + value
                        fio += value
                    if not fio: fio = None

                    userdata = UserTableData(
                        email=email,
                        login=ws.cell(row=cell.row, column=login_column).value,
                        fio=fio,
                        subjects=[],
                    )
                subject_name = ws.cell(row=cell.row, column=subject_column).value
                subject_date = ws.cell(row=cell.row, column=sel_date_column).value
                subject_date = convert_date_string(str(subject_date)) if subject_date else None
                userdata.subjects.append(TableSubject(subject_name, subject_date))

        return userdata

    def get_worksheet(self, name):
        self.check_loaded()
        for sheetname in self._xlsx.sheetnames:
            if sheetname.lower() == name.lower():
                return self._xlsx[sheetname]
        return None

    def insert_passwords(self, worksheet, first_row_is_header=True):
        """ Insert password formula into the sheet for password column 
        
            Note: use , as argument separator instead of ; in formula
                  %i for row number
            formula = ExcelDriver.get_password_formula()
        """
        formula = self.get_password_formula()
        password_column_id = self.get_column_by_name(worksheet, 'пароль')
        start_row = 2 if first_row_is_header else 1
        password_column = next(worksheet.iter_cols(min_row=start_row, 
                                min_col=password_column_id, max_col=password_column_id))
        for cell in password_column:
            cell.value = formula.replace('%i', str(cell.row))

    def load(self, filepath):
        if not os.path.isfile(filepath):
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), filepath)
        self._xlsx = load_workbook(filepath)
        self._filepath = filepath

    def remove_other_sheets(self, worksheet):
        self.check_loaded()
        for sheet in self._xlsx:
            if sheet != worksheet:
                self._xlsx.remove(sheet)

    def save(self, filepath=None):
        self.check_loaded()
        _filepath = filepath if filepath else self._filepath
        self._xlsx.save(_filepath)

    def mark_user(self, worksheet, email, first_row_is_header=True, fgColor='FF558ED5'):
        """ Mark users who should not be registered by blue color """
        fill = PatternFill('solid', fgColor=fgColor.upper())
        start_row = 2 if first_row_is_header else 1
        email_column = self.get_column_by_name(worksheet, 'email')

        for row in worksheet.iter_cols(min_col=email_column, max_col=email_column, min_row=start_row):
            for cell in row:
                if cell.value != email:
                    continue
                for cell_range in worksheet.iter_rows(min_row=cell.row, max_row=cell.row, max_col=9):
                    for cell_from_range in cell_range:
                        cell_from_range.fill = fill

    def mark_user_as_registered(self, email, first_row_is_header=True):
        self.check_loaded()
        for worksheet in self._xlsx:
            self.mark_user(worksheet, email, first_row_is_header, 'FF558ED5')
    
    def mark_user_as_skipped(self, email, first_row_is_header=True):
        self.check_loaded()
        for worksheet in self._xlsx:
            self.mark_user(worksheet, email, first_row_is_header, 'FFFF3838')

    @classmethod
    def set_comment(cls, worksheet, email, comment) -> bool:
        """ Sets the comment at the first email cell. """
        fill = PatternFill('solid', fgColor='FFA9D08E')
        email_column = cls.get_column_by_name(worksheet, 'email')
        for col_items in worksheet.iter_cols(min_col=email_column, max_col=email_column):
            for cell in col_items:
                if cell.value == email:
                    cell.comment = Comment(comment, 'elexam')
                    cell.fill = fill
                    return True
        return False
