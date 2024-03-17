import logging
import os
from typing import List, Union

import math
import openpyxl
import time
from PIL.Image import Image as PILImage
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from shouyu.config import Config
from shouyu.decorator.servicehandler import service_handler
from shouyu.service.context import ExcelContext
from shouyu.util.process import ProcessManager


class KbExcel:
    IMAGE_PATH = '../../temp.png'
    POPUP_MSG_LENGTH = 100

    def __init__(self, excel_path=None):
        self._excel_path = excel_path or Config.excel_path()
        self._worksheet_name = time.strftime('%Y-%m-%d')
        self._workbook: Workbook = self._load_workbook()
        self._changed = False
        self._pop_up_msgs = None

    def _load_workbook(self) -> Workbook:
        if not os.path.exists(self._excel_path):
            workbook = openpyxl.Workbook()
        else:
            workbook = openpyxl.load_workbook(self._excel_path)
        return workbook

    @property
    def _active_worksheet(self) -> Worksheet:
        if self._worksheet_name not in self._workbook.sheetnames:
            worksheet: Worksheet = self._workbook.create_sheet(self._worksheet_name)
            worksheet['A1'] = 'plan'
            worksheet['A7'] = 'task: '
            ExcelContext.column_steps = 1
            self._changed = True
            self._workbook.active = worksheet
        else:
            worksheet: Worksheet = self._workbook.get_sheet_by_name(self._worksheet_name)
            self._workbook.active = worksheet
        return worksheet

    def current_anchor(self, worksheet: Union[str, Worksheet]) -> str:
        if isinstance(worksheet, str):
            worksheet: Worksheet = self._workbook.get_sheet_by_name(worksheet)
        elif isinstance(worksheet, Worksheet):
            worksheet: Worksheet = worksheet
        else:
            raise RuntimeError(f'Invalid worksheet')

        max_image = self._find_max_image(worksheet)
        if max_image:
            if worksheet.max_row < max_image.anchor._from.row + math.ceil(max_image.height / 18) + 1:
                return max_image.anchor, max_image
            else:
                anchor = f'{get_column_letter(self._find_max_column_index(worksheet))}{worksheet.max_row}'
                return anchor, self._active_worksheet[anchor].value
        else:
            if worksheet.max_row == 1 and worksheet.max_column == 1 and worksheet[1][0].value is None:
                anchor = 'A1'
                return anchor, self._active_worksheet[anchor].value
            else:
                anchor = f'{get_column_letter(self._find_max_column_index(worksheet))}{worksheet.max_row}'
                return anchor, self._active_worksheet[anchor].value

    def _next_anchor(self, worksheet: Union[str, Worksheet], column_offset: int = 0, row_offset: int = 0) -> str:
        if isinstance(worksheet, str):
            worksheet: Worksheet = self._workbook.get_sheet_by_name(worksheet)
        elif isinstance(worksheet, Worksheet):
            worksheet: Worksheet = worksheet
        else:
            raise RuntimeError(f'Invalid worksheet')

        max_image = self._find_max_image(worksheet)
        if max_image:
            if worksheet.max_row < max_image.anchor._from.row + math.ceil(max_image.height / 18) + 1:
                return f'{get_column_letter(max_image.anchor._from.col + 1 + column_offset)}{max_image.anchor._from.row + math.ceil(max_image.height / 18) + 1 + row_offset}'
            else:
                return f'{get_column_letter(self._find_max_column_index(worksheet) + column_offset)}{worksheet.max_row + 1 + row_offset}'
        else:
            if worksheet.max_row == 1 and worksheet.max_column == 1 and worksheet[1][0].value is None:
                return f'A{1 + row_offset}'
            else:
                return f'{get_column_letter(self._find_max_column_index(worksheet) + column_offset)}{worksheet.max_row + 1 + row_offset}'

    @staticmethod
    def _find_max_image(sheet: Worksheet) -> Image:
        if not sheet._images:
            return None
        if len(sheet._images) == 1:
            return sheet._images[0]
        max = sheet._images[0]
        for img in sheet._images[1:]:
            if img.anchor._from.row + math.ceil(img.height / 18) > max.anchor._from.row + math.ceil(max.height / 18):
                max = img
        return max

    @staticmethod
    def _find_max_column_index(sheet: Worksheet) -> int:
        if sheet.max_column == 1:
            return 1
        row = sheet[sheet.max_row]
        for i in range(sheet.max_column, 0, -1):
            if row[i - 1].value:
                return i

        return 1

    def _append_image(self, img: PILImage, anchor: str):
        img.save(self.IMAGE_PATH)
        self._active_worksheet.add_image(Image(self.IMAGE_PATH), anchor)
        logging.info('saved image!')

    def _append_text(self, txt: str, anchor: str):
        if ExcelContext.cross_multiple_rows:
            col, row = coordinate_from_string(anchor)
            for line in txt.splitlines():
                self._active_worksheet[f'{col}{row}'] = line
                row += 1

        else:
            self._active_worksheet[anchor] = txt
        logging.info(f'saved text: {txt}!')

    @service_handler
    def append(self, data: PILImage or str or List[PILImage or str]):
        if not data:
            logging.info(f'Nothing to save!')
            return

        if isinstance(data, list):
            for record in data:
                self.append_one_record(record)
        else:
            self.append_one_record(data)

    def append_one_record(self, data):
        if not data:
            logging.info(f'not save empty or none data!')
            return

        anchor = self._next_anchor(
            self._active_worksheet,
            ExcelContext.get_column_steps_and_reset(),
            ExcelContext.get_row_steps_and_reset()
        )
        if isinstance(data, PILImage):
            self._append_image(data, anchor)
            msg = f'{anchor}: Image'
        else:
            self._append_text(data, anchor)
            msg = f'{anchor}: {str(data)}'
        self._changed = True
        self._pop_up_msgs = {
            'title': 'Submitting',
            'msg': msg[:self.POPUP_MSG_LENGTH],
            'image_path': os.path.abspath(self.IMAGE_PATH) if isinstance(data, PILImage) else None
        }

    @service_handler
    def move_column(self, step=0):
        anchor_or_image = self.current_anchor(self._active_worksheet)
        ExcelContext.column_steps += step
        old = coordinate_from_string(self._next_anchor(self._active_worksheet))[0]
        column_index = column_index_from_string(old)
        if column_index + ExcelContext.column_steps < 1:
            ExcelContext.column_steps = 1 - column_index
        if column_index + ExcelContext.column_steps > 16384:
            ExcelContext.column_steps = 16384 - column_index

        logging.info(f'move {ExcelContext.column_steps} steps')
        self._pop_up_msgs = {
            'title': self._generate_move_message(column_index, ExcelContext.column_steps),
            'msg': self._generate_status_message(anchor_or_image),
            'image_path': os.path.abspath(self.IMAGE_PATH) if isinstance(anchor_or_image[1], Image) else None
        }

    def _save_changed(self):
        if self._changed:
            try:
                self._workbook.save(self._excel_path)
            except PermissionError:
                logging.info(f'try to terminate process for {self._excel_path}')
                ProcessManager.terminate_by_path(self._excel_path)
                ExcelContext.terminated_excel = True
                self._workbook.save(self._excel_path)

    @service_handler
    def move_to_home(self):
        old = coordinate_from_string(self._next_anchor(self._active_worksheet))[0]
        column_index = column_index_from_string(old)
        ExcelContext.column_steps = 1 - column_index

        if ExcelContext.column_steps:
            logging.info(f'move {ExcelContext.column_steps} steps')
        self._pop_up_msgs = {
            'title': 'Move',
            'msg': self._generate_move_message(column_index, ExcelContext.column_steps)
        }

    @staticmethod
    def _generate_move_message(column_index: int, steps: int):
        from_column = get_column_letter(column_index)
        to_column = get_column_letter(column_index + ExcelContext.column_steps)
        mode = '' if ExcelContext.cross_multiple_rows else ' & Content in one cell'
        to_row = '' if ExcelContext.row_steps == 0 else f' & Jump {ExcelContext.row_steps} Rows'
        if steps > 0:
            return f'Move {from_column} -> {to_column}{to_row}{mode}'
        elif steps == 0:
            return f'{from_column}{to_row}{mode}'
        else:
            return f'Move {to_column} <- {from_column}{to_row}{mode}'

    @staticmethod
    def _generate_status_message(anchor_or_image):
        if isinstance(anchor_or_image[1], Image):
            anchor = anchor_or_image[0]._from
            return f'{get_column_letter(anchor.col + 1) + str(anchor.row)}: Image'
        else:
            return f'{anchor_or_image[0]}:{anchor_or_image[1]}'
