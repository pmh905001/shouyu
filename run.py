import math
import sys

from openpyxl.drawing.image import Image
from openpyxl.utils.cell import coordinate_from_string

from shouyu.log import Log
from shouyu.service.excel import KbExcel
from shouyu.util.package import Package
from shouyu.view.msgbox import MessageBox

def _build_title_from_args(argv):
    return " ".join(argv).strip()


def _calculate_current_row(anchor_or_image):
    if isinstance(anchor_or_image[1], Image):
        anchor = anchor_or_image[0]._from
        return anchor.row + math.ceil(anchor_or_image[1].height / 18)
    return coordinate_from_string(anchor_or_image[0])[1]


def main(argv):
    Package.set_cwd()
    Log.setup()

    title = _build_title_from_args(argv)
    if not title:
        MessageBox.pop_up_message("Missing Title", "请传入标题：python cmd.py \"this is a new plan title\"")
        return

    kb_excel = KbExcel()
    anchor_or_image = kb_excel.current_anchor(kb_excel._active_worksheet)
    current_row = _calculate_current_row(anchor_or_image)
    target_cell = f"A{current_row + 1}"

    kb_excel._active_worksheet[target_cell] = title
    kb_excel._changed = True
    kb_excel._pop_up_msgs = {
        "title": "Plan Updated",
        "msg": f"已写入标题: {title} -> {target_cell}",
        "image_path": None,
    }
    kb_excel._save_changed()
    MessageBox.pop_up_message(**kb_excel._pop_up_msgs)


if __name__ == '__main__':
    main(sys.argv[1:])
