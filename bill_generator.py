from collections import namedtuple
import configparser
from datetime import date
from dateutils import relativedelta
import locale
import os
import sys
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


CONFIG_NAME = "настройки.txt"
STATEMENT_COLUMNS = ['NUMBER', 'NAME', 'ACCOUNT', 'DEBT', 'DEBT_MONTHS', 'METER_LAST', 'METER_PAID']


def exception(msg):
    print(msg)
    input("Нажмите любую клавишу.")
    sys.exit(1)


class BillGenerator:
    StatementCols = namedtuple(
        'StatementCols',
        STATEMENT_COLUMNS
    )

    def __init__(self, config, statement_columns):
        self.config = config
        self.statement_columns = self.StatementCols(*statement_columns)

    def read_statement(self):

        template_filename = os.path.join(
            self.config['TEMPLATE_FOLDER'],
            self.config['TEMPLATE_FILENAME']
        )
        statement_folder = self.config['STATEMENT_FOLDER']
        first_row = int(self.config['FIRST_ROW'])
        last_row = int(self.config['LAST_ROW'])

        template_wb = load_workbook(filename=template_filename)
        bill_data = []

        statement_files = [file for file in os.listdir(statement_folder) if file.endswith(".xlsx")]
        try:
            statement_filename = statement_files[0]
        except IndexError:
            exception(f"В папке {statement_folder} нет файлов")
        statement_filename_full = os.path.join(statement_folder, statement_filename)

        statement = load_workbook(filename=statement_filename_full, data_only=True).worksheets[0]

        today = date.today()
        year = today.year
        for row_index in range(first_row, last_row + 1):
            row = statement[row_index]
            if self.is_valid(row):
                debt = float(row[self.get_col_index(self.statement_columns.DEBT)].value)
                debt_months = int(row[self.get_col_index(self.statement_columns.DEBT_MONTHS)].value)
                bill_months = self.get_bill_months(today, debt_months)

                context = {
                    '{%номер%}': row[self.get_col_index(self.statement_columns.NUMBER)].value,
                    '{%имя%}': row[self.get_col_index(self.statement_columns.NAME)].value,
                    '{%лицевой_счет%}': row[self.get_col_index(self.statement_columns.ACCOUNT)].value,
                    '{%месяц%}': bill_months,
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

    @staticmethod
    def get_bill_months(today, months):
        current_month = today.strftime('%B')
        if months <= 0:
            return current_month

        first_month = (today - relativedelta(months=months)).strftime('%B')
        return '%s - %s' % (first_month, current_month)

    def fill_template(self, template_wb, context):
        sheet = template_wb.worksheets[0]

        output_filename = self.config['OUTPUT_FILENAME_FORMAT']
        output_folder = self.config['OUTPUT_FOLDER']
        for row in sheet:
            for cell in row:
                for key, value in context.items():
                    if key in str(cell.value):
                        cell.value = cell.value.replace(key, str(value))
                    output_filename = output_filename.replace(key, str(value))
        output_filename_full = os.path.join(output_folder, output_filename)
        template_wb.save(filename=output_filename_full)

    def is_valid(self, row):
        required_cols = [
            self.statement_columns.NUMBER,
            self.statement_columns.NAME,
            self.statement_columns.DEBT_MONTHS,
        ]
        debt_months = int(self.config['DEBT_MONTHS'])
        if all([bool(row[self.get_col_index(col)].value) for col in required_cols]) \
                and row[self.get_col_index(self.statement_columns.DEBT_MONTHS)].value >= debt_months:
            return True
        return False


class App:
    def __init__(self):
        self.exit_flag = False
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        if os.path.exists(CONFIG_NAME):
            self.config, self.statement_columns = self.read_config()
        else:
            self.config = self.generate_default_config()
            self.exit_flag = True
        folders_list = [
            'STATEMENT_FOLDER',
            'OUTPUT_FOLDER',
            'TEMPLATE_FOLDER',
        ]
        for folder in folders_list:
            if not os.path.exists(self.config[folder]):
                os.mkdir(self.config[folder])
                self.exit_flag = True

    @staticmethod
    def generate_default_config():
        config = configparser.RawConfigParser(allow_no_value=True)
        config.optionxform = str
        config.add_section('SETTINGS')
        config.set('SETTINGS', '# Папка с ведомостью')
        config.set('SETTINGS', 'STATEMENT_FOLDER', 'Ведомость')
        config.set('SETTINGS', '')
        config.set('SETTINGS', '# Папка с шаблоном квитанции и имя самого файла-шаблона')
        config.set('SETTINGS', 'TEMPLATE_FOLDER', 'Шаблон')
        config.set('SETTINGS', 'TEMPLATE_FILENAME', 'квитанция.xlsx')
        config.set('SETTINGS', ' ')
        config.set('SETTINGS', '# Папка, куда будут сложены все квитанции')
        config.set('SETTINGS', 'OUTPUT_FOLDER', 'Квитанции')
        config.set('SETTINGS', '  ')
        config.set('SETTINGS', '# Формат имени выходного файла')
        config.set('SETTINGS', 'OUTPUT_FILENAME_FORMAT', '{%номер%}_{%месяц%}_{%имя%}.xlsx')
        config.set('SETTINGS', '   ')
        config.set('SETTINGS', '# Первый и последний ряды в ведомости')
        config.set('SETTINGS', 'FIRST_ROW', '9')
        config.set('SETTINGS', 'LAST_ROW', '170')
        config.set('SETTINGS', '    ')
        config.set('SETTINGS', '# Количество месяцев просроченных платежей')
        config.set('SETTINGS', 'DEBT_MONTHS', '3')

        config.add_section('COLUMNS')
        config.set('COLUMNS', '# Столбцы')
        config.set('COLUMNS', '# Номер по порядку')
        config.set('COLUMNS', 'NUMBER', 'A')
        config.set('COLUMNS', '# Имя')
        config.set('COLUMNS', 'NAME', 'C')
        config.set('COLUMNS', '# Номер лицевого счета')
        config.set('COLUMNS', 'ACCOUNT', 'D')
        config.set('COLUMNS', '# Долг')
        config.set('COLUMNS', 'DEBT', 'I')
        config.set('COLUMNS', '# Месяцев задолжность')
        config.set('COLUMNS', 'DEBT_MONTHS', 'J')
        config.set('COLUMNS', '# Последнее показание счетчика')
        config.set('COLUMNS', 'METER_LAST', 'K')
        config.set('COLUMNS', '# Предыдущее оплаченное')
        config.set('COLUMNS', 'METER_PAID', 'L')

        with open(CONFIG_NAME, 'w', encoding='utf-8') as file:
            config.write(file)

        return config['SETTINGS']

    @staticmethod
    def read_config():
        config = configparser.RawConfigParser()
        config.read(CONFIG_NAME, encoding="utf-8")

        cols_conf = config['COLUMNS']

        statement_columns = [cols_conf[col] for col in STATEMENT_COLUMNS]

        return config['SETTINGS'], statement_columns

    def run(self):
        try:
            bill_generator = BillGenerator(self.config, self.statement_columns)
            template_wb, bill_data = bill_generator.read_statement()
            for context in bill_data:
                bill_generator.fill_template(template_wb, context)
        except PermissionError:
            exception("Произошла ошибка. Закройте все файлы Excel перед запуском.")
        except Exception:
            exception("Произошла непредвиденная ошибка.")


if __name__ == '__main__':
    app = App()
    if not app.exit_flag:
        app.run()
    else:
        exception("Приложение инициализировано. "
                  "Вложите файлы в соответствующие папки и запустите еще раз.")
