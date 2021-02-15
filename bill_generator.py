import configparser
from datetime import date
import locale
import os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


CONFIG_NAME = "настройки.txt"


class BillGenerator:
    class StatementCols:
        NUMBER = 'A'
        NAME = 'C'
        ACCOUNT = 'D'
        DEBT = 'I'

    def __init__(self, config):
        self.config = config

    def read_statement(self):

        template_filename = os.path.join(
            self.config['STRUCTURE']['TEMPLATE_FOLDER'],
            self.config['STRUCTURE']['TEMPLATE_FILENAME']
        )
        statement_folder = self.config['STRUCTURE']['STATEMENT_FOLDER']
        first_row = int(self.config['STRUCTURE']['FIRST_ROW'])
        last_row = int(self.config['STRUCTURE']['LAST_ROW'])

        template_wb = load_workbook(filename=template_filename)
        bill_data = []

        statement_files = [file for file in os.listdir(statement_folder) if file.endswith(".xlsx")]
        try:
            statement_filename = statement_files[0]
        except IndexError:
            print(f"В папке {statement_folder} нет файлов")
            exit()
        statement_filename_full = os.path.join(statement_folder, statement_filename)

        statement = load_workbook(filename=statement_filename_full, data_only=True).worksheets[0]

        today = date.today()
        month = today.strftime('%B')
        year = today.year
        for row_index in range(first_row, last_row + 1):
            row = statement[row_index]
            if self.is_valid(row):
                debt = float(row[self.get_col_index(self.StatementCols.DEBT)].value)

                context = {
                    '{%номер%}': row[self.get_col_index(self.StatementCols.NUMBER)].value,
                    '{%имя%}': row[self.get_col_index(self.StatementCols.NAME)].value,
                    '{%лицевой_счет%}': row[self.get_col_index(self.StatementCols.ACCOUNT)].value,
                    '{%месяц%}': month,
                    '{%год%}': year,
                    '{%долг%}': ("%.2f" % debt).replace('.', ','),
                    '{%долг_рубли%}': "%.f" % debt,
                    '{%долг_копейки%}': '0',
                }
                bill_data.append(context)
        return template_wb, bill_data

    @staticmethod
    def get_col_index(letter):
        return column_index_from_string(letter) - 1

    def fill_template(self, template_wb, context):
        sheet = template_wb.worksheets[0]

        output_filename = self.config['STRUCTURE']['OUTPUT_FILENAME_FORMAT']
        output_folder = self.config['STRUCTURE']['OUTPUT_FOLDER']
        for row in sheet:
            for cell in row:
                for key, value in context.items():
                    if key in str(cell.value):
                        cell.value = cell.value.replace(key, str(value))
                    output_filename = output_filename.replace(key, str(value))
        output_filename_full = os.path.join(output_folder, output_filename)
        template_wb.save(filename=output_filename_full)

    @classmethod
    def is_valid(cls, row):
        if all([bool(row[cls.get_col_index(col)].value) for col in [
                    cls.StatementCols.NUMBER,
                    cls.StatementCols.NAME,
                ]]):
            return True
        return False


class App:
    def __init__(self):
        self.exit_flag = False
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        if os.path.exists(CONFIG_NAME):
            self.config = self.read_config()
        else:
            self.config = self.generate_default_config()
            self.exit_flag = True
        folders_list = [
            'STATEMENT_FOLDER',
            'OUTPUT_FOLDER',
            'TEMPLATE_FOLDER',
        ]
        for folder in folders_list:
            if not os.path.exists(self.config['STRUCTURE'][folder]):
                os.mkdir(self.config['STRUCTURE'][folder])
                self.exit_flag = True

    @staticmethod
    def generate_default_config():
        config = configparser.RawConfigParser(allow_no_value=True)
        config.optionxform = str
        config.add_section('STRUCTURE')
        config.set('STRUCTURE', '# Папка с ведомостью')
        config.set('STRUCTURE', 'STATEMENT_FOLDER', 'Ведомость')
        config.set('STRUCTURE', '')
        config.set('STRUCTURE', '# Папка с шаблоном квитанции и имя самого файла-шаблона')
        config.set('STRUCTURE', 'TEMPLATE_FOLDER', 'Шаблон')
        config.set('STRUCTURE', 'TEMPLATE_FILENAME', 'квитанция.xlsx')
        config.set('STRUCTURE', ' ')
        config.set('STRUCTURE', '# Папка, куда будут сложены все квитанции')
        config.set('STRUCTURE', 'OUTPUT_FOLDER', 'Квитанции')
        config.set('STRUCTURE', '  ')
        config.set('STRUCTURE', '# Формат имени выходного файла')
        config.set('STRUCTURE', 'OUTPUT_FILENAME_FORMAT', '{%номер%}_{%месяц%}_{%имя%}.xlsx')
        config.set('STRUCTURE', '   ')
        config.set('STRUCTURE', '# Первый и последний ряды в ведомости')
        config.set('STRUCTURE', 'FIRST_ROW', '9')
        config.set('STRUCTURE', 'LAST_ROW', '170')

        config.add_section('COLUMNS')
        config.set('COLUMNS', '# Столбцы')
        config.set('COLUMNS', 'NUMBER', 'A')
        config.set('COLUMNS', 'NAME', 'C')
        config.set('COLUMNS', 'ACCOUNT', 'D')
        config.set('COLUMNS', 'DEBT', 'I')

        with open(CONFIG_NAME, 'w', encoding='utf-8') as file:
            config.write(file)

        return config

    @staticmethod
    def read_config():
        config = configparser.RawConfigParser()
        config.read(CONFIG_NAME, encoding="utf-8")

        cols_conf = config['COLUMNS']
        BillGenerator.StatementCols.NAME = cols_conf['NAME']
        BillGenerator.StatementCols.NUMBER = cols_conf['NUMBER']
        BillGenerator.StatementCols.DEBT = cols_conf['DEBT']
        BillGenerator.StatementCols.ACCOUNT = cols_conf['ACCOUNT']

        return config

    def run(self):
        try:
            bill_generator = BillGenerator(self.config)
            template_wb, bill_data = bill_generator.read_statement()
            for context in bill_data:
                bill_generator.fill_template(template_wb, context)
        except PermissionError:
            print("Произошла ошибка. Закройте все файлы Excel перед запуском.")
        except Exception:
            print("Произошла непредвиденная ошибка.")
            raise


if __name__ == '__main__':
    app = App()
    if not app.exit_flag:
        app.run()
    else:
        print("Приложение инициализировано. Вложите файлы в соответствующие папки и запустите еще раз.")
