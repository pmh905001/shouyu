class ExcelContext:
    column_steps = 0
    row_steps = 0
    one_cell_mode = True

    @classmethod
    def get_column_steps_and_reset(cls):
        tmp = cls.column_steps
        cls.column_steps = 0
        return tmp

    @classmethod
    def get_row_steps_and_reset(cls):
        tmp = cls.row_steps
        cls.row_steps = 0
        return tmp
