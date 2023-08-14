import logging
import math
import os
import time
from typing import Union

import openpyxl
from PIL.Image import Image as PILImage
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_context import ExcelContext
from msg_box import MessageBox, MessageType
from process import ProcessManager


class ExcelWriter:

    def __init__(self, excel_path):
        self._excel_path = excel_path
        self._worksheet_name = time.strftime('%Y-%m-%d')
        self._workbook: Workbook = self._load_workbook()

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
            self._workbook.active = worksheet
        else:
            worksheet: Worksheet = self._workbook.get_sheet_by_name(self._worksheet_name)
            self._workbook.active = worksheet
        return worksheet

    def _next_anchor(self, worksheet: Union[str, Worksheet], column_offset: int = 0) -> str:
        if isinstance(worksheet, str):
            worksheet: Worksheet = self._workbook.get_sheet_by_name(worksheet)
        elif isinstance(worksheet, Worksheet):
            worksheet: Worksheet = worksheet
        else:
            raise RuntimeError(f'Invalid worksheet')

        max_image = self._find_max_image(worksheet)
        if max_image:
            if worksheet.max_row < max_image.anchor._from.row + math.ceil(max_image.height / 18) + 1:
                return f'{get_column_letter(max_image.anchor._from.col + 1 + column_offset)}{max_image.anchor._from.row + math.ceil(max_image.height / 18) + 1}'
            else:
                return f'{get_column_letter(self._find_max_column_index(worksheet) + column_offset)}{worksheet.max_row + 1}'
        else:
            if worksheet.max_row == 1 and worksheet.max_column == 1 and worksheet[1][0].value is None:
                return 'A1'
            else:
                return f'{get_column_letter(self._find_max_column_index(worksheet) + column_offset)}{worksheet.max_row + 1}'

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

    def _save_image(self, img: PILImage):
        anchor = self._next_anchor(self._active_worksheet, ExcelContext.get_steps_and_reset())
        tmp_img = 'temp.png'
        img.save(tmp_img)
        self._active_worksheet.add_image(Image(tmp_img), anchor)
        self._workbook.save(self._excel_path)
        logging.info('saved image!')

    def _save_text(self, txt: str):
        anchor = self._next_anchor(self._active_worksheet, ExcelContext.get_steps_and_reset())
        self._active_worksheet[anchor] = txt
        self._workbook.save(self._excel_path)
        logging.info(f'saved text: {txt}!')

    def save(self, data: PILImage or str, ignore_permission_error: bool = False):
        try:
            if data:
                if isinstance(data, PILImage):
                    self._save_image(data)
                else:
                    self._save_text(data)
                MessageBox.pop_up_message('Success', str(data), MessageType.SUCCESS)
                # Tray._icon.notify(str(data), 'Success')
            else:
                logging.info(f'Nothing to save!')
        except PermissionError as permission_ex:
            if not ignore_permission_error:
                logging.info(f'try to terminate process for {self._excel_path}')
                ProcessManager.terminate_by_path(self._excel_path)
                try:
                    self._workbook.save(self._excel_path)
                    MessageBox.pop_up_message('Success', str(data), MessageType.SUCCESS)
                except Exception as ex:
                    logging.exception(f'failed to save "{data}"')
                    MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR)

                ProcessManager.resume_last_closed_process(self._excel_path)
            else:
                logging.exception(f'failed to save "{data}"')
                MessageBox.pop_up_message('Failed', str(permission_ex), MessageType.ERROR)
        except Exception as ex:
            logging.exception(f'failed to save "{data}"')
            MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR)

    def move_column(self, step=0):
        ExcelContext.steps += step
        old = coordinate_from_string(self._next_anchor(self._active_worksheet))[0]
        column_index = column_index_from_string(old)
        if column_index + ExcelContext.steps < 1:
            ExcelContext.steps = 1 - column_index

        logging.info(f'move {ExcelContext.steps} steps')
        MessageBox.pop_up_message('Move', self._generate_move_message(column_index,ExcelContext.steps), MessageType.SUCCESS)

    @staticmethod
    def _generate_move_message(column_index: int, steps: int):
        from_column = get_column_letter(column_index)
        to_column = get_column_letter(column_index + ExcelContext.steps)
        if steps > 0:
            return f'{from_column} -> {to_column}'
        elif steps == 0:
            return from_column
        else:
            return f'{to_column} <- {from_column}'
