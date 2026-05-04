import glob
import logging
import os
import shutil
from io import BytesIO
from typing import List, Optional, Union
from zipfile import BadZipFile

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
from shouyu.service.plan import PlanService
from shouyu.util.process import ProcessManager


class KbExcel:
    IMAGE_PATH = '../../temp.png'
    POPUP_MSG_LENGTH = 100

    def __init__(self, excel_path=None):
        self._excel_path = excel_path or Config.excel_path()
        self._worksheet_name = time.strftime('%Y-%m-%d')
        # Track whether we had to recover from backup so the caller can warn UX-side.
        self._recovered_from_backup: Optional[str] = None
        self._workbook: Workbook = self._load_workbook()
        self._changed = False
        self._pop_up_msgs = None

    def _load_workbook(self) -> Workbook:
        if not os.path.exists(self._excel_path):
            return openpyxl.Workbook()

        try:
            workbook = openpyxl.load_workbook(self._excel_path)
        except (KeyError, BadZipFile, ValueError) as e:
            # Common corruption signature is `KeyError: '[Content_Types].xml'`
            # caused by a previous in-place save being interrupted. Try to
            # recover automatically from the most recent good backup so the
            # user is never locked out of their own data.
            logging.error(
                f"main Excel file looks corrupt ({e}); attempting backup recovery"
            )
            workbook = self._recover_from_backup()
            if workbook is None:
                logging.warning(
                    "no usable backup found; starting from a fresh empty workbook"
                )
                workbook = openpyxl.Workbook()

        # See `_materialize_images` docstring for why this is critical.
        self._materialize_images(workbook)
        return workbook

    def _recover_from_backup(self) -> Optional[Workbook]:
        """Walk backups newest-first, return the first one that opens cleanly."""
        for backup in self.list_backups(self._excel_path):
            try:
                wb = openpyxl.load_workbook(backup)
                self._recovered_from_backup = backup
                logging.warning(f"auto-recovered workbook from backup: {backup}")
                return wb
            except Exception:
                logging.warning(f"backup is also unreadable, trying next: {backup}")
                continue
        return None

    @staticmethod
    def _materialize_images(workbook: Workbook) -> None:
        """Copy every image's bytes into an in-memory BytesIO right after load.

        openpyxl stores image refs as `ZipExtFile` objects bound to the
        workbook's underlying zip archive. That archive is closed (or GC'd)
        at unpredictable times. The next save then explodes with
        `ValueError: I/O operation on closed file.` partway through writing
        the new xlsx — leaving the destination half-written and corrupt.

        Materializing once on load decouples image data from the loaded zip
        handle, so subsequent saves are safe.
        """
        for ws in workbook.worksheets:
            images = getattr(ws, "_images", None) or []
            for img in list(images):
                try:
                    ref = getattr(img, "ref", None)
                    if ref is None:
                        continue
                    if isinstance(ref, (str, bytes, BytesIO)):
                        continue
                    if hasattr(ref, "read"):
                        try:
                            ref.seek(0)
                        except Exception:
                            pass
                        data = ref.read()
                        if data:
                            img.ref = BytesIO(data)
                except Exception:
                    logging.exception("failed to materialize image into memory")

    @property
    def recovered_from_backup(self) -> Optional[str]:
        """Path of the backup we auto-recovered from, or None for a normal load."""
        return self._recovered_from_backup

    @property
    def _active_worksheet(self) -> Worksheet:
        if "todo list" in  self._workbook.sheetnames and self._workbook.sheetnames[len(self._workbook.sheetnames)-1] != "todo list":
           worksheet: Worksheet = self._workbook.get_sheet_by_name("todo list") 
           self._workbook.move_sheet(worksheet,offset=len(self._workbook.sheetnames)-1)
           self._changed = True
           
        
        if self._worksheet_name not in self._workbook.sheetnames:
            worksheet: Worksheet = self._workbook.create_sheet(self._worksheet_name)
            PlanService(worksheet).seed_default_plan()
            self._changed = True
            self._workbook.active = worksheet
        else:
            worksheet: Worksheet = self._workbook.get_sheet_by_name(self._worksheet_name)
            self._workbook.active = worksheet

        for ws in self._workbook.worksheets:
            expected = (ws == worksheet)
            if ws.views.sheetView[0].tabSelected != expected:
                ws.views.sheetView[0].tabSelected = expected
                self._changed = True

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
        self._active_worksheet[anchor]=f'Image created at {time.strftime("%Y-%m-%d %H:%M:%S")}'
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

        target_col = ExcelContext.get_target_column_and_reset()

        if isinstance(data, (list, tuple)):
            for record in data:
                self.append_one_record(record, target_col)
        else:
            self.append_one_record(data, target_col)

    def append_one_record(self, data, target_column=None):
        if not data:
            logging.info(f'not save empty or none data!')
            return

        anchor = self._next_anchor(
            self._active_worksheet,
            ExcelContext.get_column_steps_and_reset(),
            ExcelContext.get_row_steps_and_reset()
        )
        if target_column:
            _, row = coordinate_from_string(anchor)
            anchor = f'{target_column}{row}'

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

    def _backup_excel(self):
        if not os.path.exists(self._excel_path):
            return

        excel_dir = os.path.dirname(self._excel_path) or '.'
        name, ext = os.path.splitext(os.path.basename(self._excel_path))
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        backup_path = os.path.join(excel_dir, f'{name}_backup_{timestamp}{ext}')
        shutil.copy2(self._excel_path, backup_path)
        logging.info(f'backup created: {backup_path}')

        pattern = os.path.join(excel_dir, f'{name}_backup_*{ext}')
        backups = sorted(glob.glob(pattern), key=os.path.getmtime)
        max_backups = Config.max_backups()
        while len(backups) > max_backups:
            oldest = backups.pop(0)
            os.remove(oldest)
            logging.info(f'removed old backup: {oldest}')

    def _save_changed(self):
        if not self._changed:
            return
        self._backup_excel()
        self._atomic_save()
        # If we recovered from backup and just persisted the recovered workbook
        # back to the canonical path, we're back in sync — clear the flag so
        # subsequent saves don't keep flagging recovery in stats.
        if self._recovered_from_backup is not None:
            logging.info(
                f"persisted recovered workbook back to {self._excel_path}; "
                f"original recovery source was {self._recovered_from_backup}"
            )
            self._recovered_from_backup = None

    def _atomic_save(self) -> None:
        """Save to a temp file then atomically replace the original.

        If save fails midway (PermissionError, the openpyxl image bug, anything),
        the original file stays untouched. This is the single most important
        safety net — without it a partial save corrupts the canonical Excel.
        """
        excel_dir = os.path.dirname(self._excel_path) or '.'
        name, ext = os.path.splitext(os.path.basename(self._excel_path))
        tmp_path = os.path.join(excel_dir, f'.{name}.tmp_{os.getpid()}{ext}')
        try:
            self._workbook.save(tmp_path)
        except Exception:
            # Clean up the half-written tmp file before re-raising.
            self._safe_remove(tmp_path)
            raise
        # save succeeded; flip the canonical file in one OS call.
        try:
            os.replace(tmp_path, self._excel_path)
        except PermissionError:
            # The canonical file is currently locked (most often: the user has
            # opened it in MS Excel / WPS). The new content is sitting safely
            # in tmp_path — we keep it on disk so the user can recover it.
            kept = os.path.join(
                excel_dir,
                f'{name}.unsaved_{time.strftime("%Y%m%d_%H%M%S")}{ext}',
            )
            try:
                os.replace(tmp_path, kept)
                logging.error(
                    f"target Excel is locked; pending changes preserved at: {kept}"
                )
            except Exception:
                self._safe_remove(tmp_path)
                logging.exception("failed to preserve unsaved changes")
            raise

    @staticmethod
    def _safe_remove(path: str) -> None:
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            logging.exception(f"failed to remove temp file: {path}")

    @staticmethod
    def list_backups(excel_path: str) -> List[str]:
        """Return all `*_backup_*.xlsx` files for `excel_path`, newest first."""
        if not excel_path:
            return []
        excel_dir = os.path.dirname(excel_path) or '.'
        name, ext = os.path.splitext(os.path.basename(excel_path))
        pattern = os.path.join(excel_dir, f'{name}_backup_*{ext}')
        try:
            return sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)
        except Exception:
            logging.exception("failed to list backups")
            return []

    @classmethod
    def restore_from_backup(cls, backup_path: str, excel_path: Optional[str] = None) -> str:
        """Restore the canonical Excel file from `backup_path`.

        Before overwriting, the *current* canonical file (which the user might
        still want — e.g. if they picked the wrong backup) is itself backed up
        with a `pre_restore_<ts>` suffix so this operation is reversible.

        Returns the pre-restore safety-backup path (empty string if none).
        Raises on failure.
        """
        target = excel_path or Config.excel_path()
        if not backup_path or not os.path.isfile(backup_path):
            raise FileNotFoundError(f"backup not found: {backup_path}")

        # Validate the backup actually opens before we touch anything.
        try:
            openpyxl.load_workbook(backup_path).close()
        except Exception as e:
            raise RuntimeError(f"backup file is itself unreadable: {e}") from e

        safety_path = ''
        if os.path.exists(target):
            target_dir = os.path.dirname(target) or '.'
            tname, text = os.path.splitext(os.path.basename(target))
            safety_path = os.path.join(
                target_dir,
                f'{tname}.pre_restore_{time.strftime("%Y%m%d_%H%M%S")}{text}',
            )
            try:
                shutil.copy2(target, safety_path)
                logging.info(f"pre-restore safety copy saved to: {safety_path}")
            except Exception:
                logging.exception("failed to take pre-restore safety copy")
                safety_path = ''

        shutil.copy2(backup_path, target)
        logging.info(f"restored Excel from backup: {backup_path} -> {target}")
        return safety_path

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

    @classmethod
    def append_title_to_next_row(cls, title: str):
        instance = cls()
        anchor_or_image = instance.current_anchor(instance._active_worksheet)
        if isinstance(anchor_or_image[1], Image):
            anchor = anchor_or_image[0]._from
            current_row = anchor.row + math.ceil(anchor_or_image[1].height / 18)
        else:
            _, current_row = coordinate_from_string(anchor_or_image[0])
        target_cell = f"A{current_row + 1}"
        instance._active_worksheet[target_cell] = title
        instance._changed = True
        instance._pop_up_msgs = {
            "title": "Plan Updated",
            "msg": f"已写入标题: {title} -> {target_cell}",
            "image_path": None,
        }
        instance._save_changed()
        from shouyu.view.msgbox import MessageBox
        MessageBox.pop_up_message(**instance._pop_up_msgs)

    def plan_service(self) -> PlanService:
        return PlanService(self._active_worksheet)

    def plan_service_for(self, date_str: str):
        """Return a PlanService bound to <date_str>'s worksheet, or None if it doesn't exist."""
        if not date_str or date_str not in self._workbook.sheetnames:
            return None
        return PlanService(self._workbook[date_str])

    def write_reflection(self, text: str, date_str: str = None) -> None:
        date_str = date_str or time.strftime('%Y-%m-%d')
        if 'reflections' in self._workbook.sheetnames:
            ws = self._workbook['reflections']
        else:
            ws = self._workbook.create_sheet('reflections')
            ws['A1'] = 'date'
            ws['B1'] = 'reflection'
        target_row = None
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == date_str:
                target_row = row
                break
        if target_row is None:
            target_row = max(ws.max_row, 1) + 1
            ws.cell(row=target_row, column=1, value=date_str)
        ws.cell(row=target_row, column=2, value=text or '')
        self._changed = True
        self._save_changed()

    def read_reflection(self, date_str: str = None) -> str:
        date_str = date_str or time.strftime('%Y-%m-%d')
        if 'reflections' not in self._workbook.sheetnames:
            return ''
        ws = self._workbook['reflections']
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == date_str:
                return str(ws.cell(row=row, column=2).value or '')
        return ''

    def append_detail(self, text: str, column: str = 'B') -> str:
        """Append a single detail value into the active area at the next available row.

        Used by pomodoro and other auto-loggers that should not pop up message boxes.
        Returns the anchor written to.
        """
        if not text:
            return ''
        plan = self.plan_service()
        last_used = plan.last_used_row_in_active_area()
        start = plan.active_area_start_row()
        target_row = max(last_used + 1, start)
        anchor = f'{column}{target_row}'
        self._active_worksheet[anchor] = text
        self._changed = True
        self._save_changed()
        return anchor

    def force_save(self) -> None:
        """Public helper to flush pending changes regardless of decorator pipeline."""
        self._save_changed()

    def mark_changed(self) -> None:
        self._changed = True

    @property
    def workbook(self) -> Workbook:
        return self._workbook

    @property
    def active_worksheet(self) -> Worksheet:
        return self._active_worksheet
