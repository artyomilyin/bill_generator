import os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


STATEMENT_FOLDER = 'Ведомость'

TEMPLATE_FOLDER = 'Шаблон'
TEMPLATE_FILENAME = 'квитанция.xlsx'

OUTPUT_FOLDER = 'Квитанции'
OUTPUT_FILENAME_FORMAT = '{%номер%}_{%месяц%}_{%имя%}.xlsx'

FIRST_ROW = 9




class BillGenerator:
    class StatementCols:
        NUMBER = 'A'
        NAME = 'C'
        ACCOUNT = 'D'
        DEBT = 'I'

    @staticmethod
    def get_col_index(letter):
        return column_index_from_string(letter) - 1

    @staticmethod
    def fill_template(context, wb):
        sheet = wb.worksheets[0]

        output_filename = OUTPUT_FILENAME_FORMAT
        for row in sheet:
            for cell in row:
                for key, value in context.items():
                    if key in str(cell.value):
                        cell.value = str(cell.value.replace(key, value))
                    output_filename = output_filename.replace(key, value)
        output_filename_full = os.path.join(OUTPUT_FOLDER, output_filename)
        wb.save(filename=output_filename_full)

    @classmethod
    def is_valid(cls, row):
        if all([bool(row[cls.get_col_index(col)].value) for col in [
                    cls.StatementCols.NUMBER,
                    cls.StatementCols.NAME,
                    cls.StatementCols.ACCOUNT,
                    cls.StatementCols.DEBT,
                ]]):
            return True
        return False


    @classmethod
    def read_statement(cls):

        statement_files = [file for file in os.listdir(STATEMENT_FOLDER) if file.endswith(".xls") or file.endswith(".xlsx")]
        try:
            statement_filename = statement_files[0]
        except IndexError:
            print(f"В папке {STATEMENT_FOLDER} нет файлов")
            exit()
        statement_filename_full = os.path.join(STATEMENT_FOLDER, statement_filename)

        statement = load_workbook(filename=statement_filename_full, data_only=True).worksheets[0]
        result = []
        for row in statement:
            if cls.is_valid(row):
                print(row)

                context = {
                    '{%номер%}': row[cls.get_col_index(cls.StatementCols.NUMBER)].value,
                    '{%имя%}': row[cls.get_col_index(cls.StatementCols.NAME)].value,
                    '{%лицевой_счет%}': row[cls.get_col_index(cls.StatementCols.ACCOUNT)].value,
                    '{%месяц%}': 'Февраль',
                    '{%год%}': '2021',
                    '{%долг%}': row[cls.get_col_index(cls.StatementCols.DEBT)].value,
                    '{%долг_рубли%}': '4043',
                    '{%долг_копейки%}': '00',
                }
                print(context['{%номер%}'])
                result.append(context)

        return result


if __name__ == '__main__':
    try:
        bill_generator = BillGenerator()
        result = bill_generator.read_statement()

        template_filename = os.path.join(TEMPLATE_FOLDER, TEMPLATE_FILENAME)
        template_wb = load_workbook(filename=template_filename)
        for context in result:
            bill_generator.fill_template(context, template_wb)
    except PermissionError:
        print("Произошла ошибка. Закройте все файлы Excel перед запуском.")
    except:
        print("Произошла непредвиденная ошибка.")
        raise