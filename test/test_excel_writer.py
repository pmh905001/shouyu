import pytest

from excel_writer import ExcelWriter


@pytest.mark.parametrize(
    'worksheet_name,expect_anchor',
    [
        ('nothing', 'A1'),
        ('1_row', 'A2'),
        ('1_row_B', 'B2'),
        ('B2', 'B3'),
        ('one_image', 'A2'),
        ('2_images', 'A3'),
        ('one_image_on_b_column', 'B2'),
        ('one_image_on_row_2', 'B3'),
        ('text_and_image', 'B3'),
        ('image_and_text', 'B3'),
    ]
)
def test_next_anchor(worksheet_name, expect_anchor):
    writer = ExcelWriter('calculate_next_position.xlsx')
    assert writer._next_anchor(worksheet_name) == expect_anchor
