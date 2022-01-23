import sqlite3
import xlrd
from datetime import date, timedelta


datatypes = {1: "TEXT", 2: "DOUBLE", 3: "TEXT"}

print("XLS to SQL converter v0.2 by Ph.")

def add_quotes(string):
    # костыль. честно.
    if '"' not in str(string):
        return '"' + str(string) + '"'
    else:
        return string

class MySQLBase():
    def __init__(self, db_filename):
        self.con = sqlite3.connect(db_filename)
        self.cur = self.con.cursor()  # init database cursor

    def execute(self, request):
        # execute request
        try:
            self.cur.execute(request)
            self.con.commit()
        except sqlite3.OperationalError as e:
            print(f"WARNING! {e}")
            print(request)

    def add_table(self, name, columns):
        # types: 1 - str, 2 - float, 3 - date
        new_request = f"""CREATE TABLE IF NOT EXISTS "{name}" ("""
        is_primary_key = True
        for column in columns.keys():
            new_request += f'"{column}" {datatypes[columns[column]]}'
            if is_primary_key:
                # adding primary key to the first column
                new_request += " PRIMARY KEY"
                is_primary_key = False
            new_request += ",\n"
        new_request = new_request[:-2] + ");"  # ending request

        self.execute(new_request)

    def add_row(self, table_name, data):
        self.execute(f'INSERT OR IGNORE INTO "{table_name}" VALUES ({", ".join(map(lambda x: add_quotes(x), data))})')

    def close(self):
        self.cur.close()
        self.con.close()


def main():
    filename = input("Enter XLS file name: ")
    basename = input("Enter new base name: ")
    xlsfile = xlrd.open_workbook_xls(filename=filename, formatting_info=False)
    sheets = xlsfile.sheets()
    base = MySQLBase(basename)
    for s_id, sheet in enumerate(sheets):
        rows = list(sheet.get_rows())
        print(f"Adding sheet {s_id+1}/{len(sheets)}")
        new_table_columns = dict()
        for col_i in range(sheet.ncols):
            cell = rows[1][col_i]
            if cell.ctype == 3:
                date_to_text = (date(1900, 1, 1) + timedelta(int(cell.value))).strftime('%d.%m.%Y')
                cell.ctype = 1
                cell.value = date_to_text

            new_table_columns[rows[0][col_i].value] = cell.ctype
        base.add_table(sheet.name, new_table_columns)

        for row_i in range(1, sheet.nrows):
            row = rows[row_i]
            new_row = []
            for cell in row:
                # convert date if exists
                if cell.ctype == 3:
                    date_to_text = (date(1900, 1, 1) + timedelta(int(cell.value))).strftime('%d.%m.%Y')
                    cell.ctype = 1
                    cell.value = date_to_text
                new_row.append(cell)
            base.add_row(sheet.name, map(lambda x: x.value, new_row))


    # base.cur.execute('SELECT * FROM "Магазин"')
    # print(base.cur.fetchone())
    print("Database saved.")
    base.close()


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
