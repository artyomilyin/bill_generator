import os
from openpyxl import load_workbook


TEMPLATE_FOLDER = 'Шаблон'
TEMPLATE_FILENAME = 'квитанция.xlsx'

OUTPUT_FOLDER = 'Квитанции'
OUTPUT_FILENAME_FORMAT = '{%номер%}_{%месяц%}_{%имя%}.xlsx'

filename = os.path.join(TEMPLATE_FOLDER, TEMPLATE_FILENAME)


def fill_template(context, template_sheet):
    output_filename = OUTPUT_FILENAME_FORMAT
    for row in template_sheet:
        for cell in row:
            for key, value in context.items():
                if key in str(cell.value):
                    cell.value = str(cell.value.replace(key, value))
                output_filename = output_filename.replace(key, value)
    output_filename_full = os.path.join(OUTPUT_FOLDER, output_filename)
    wb.save(filename=output_filename_full)

wb = load_workbook(filename=filename)

sheet = wb.worksheets[0]
context = {
    '{%номер%}': '123',
    '{%имя%}': 'Иванов Иван Сергеевич',
    '{%лицевой_счет%}': '123-123',
    '{%месяц%}': 'Февраль',
    '{%год%}': '2021',
    '{%долг%}': '4043,00',
    '{%долг_рубли%}': '4043',
    '{%долг_копейки%}': '00',
}
fill_template(context, sheet)

