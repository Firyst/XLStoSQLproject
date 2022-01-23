import sqlite3
import xlrd

filename = "3-0.xls"

datatypes = {1: "TEXT", 2: "DOUBLE", 3: "TEXT"}

print("XLS to SQL converter v0.1 by Ph.")

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
    xlsfile = xlrd.open_workbook_xls(filename=filename, formatting_info=False)
    sheets = xlsfile.sheets()
    print(len(sheets))
    base = MySQLBase("test.db")
    for s_id, sheet in enumerate(sheets):
        rows = list(sheet.get_rows())
        print(f"Adding sheet {s_id+1}/{len(sheets)}")
        new_table_columns = dict()
        for col_i in range(sheet.ncols):
            new_table_columns[rows[0][col_i].value] = rows[1][col_i].ctype
        base.add_table(sheet.name, new_table_columns)

        for row_i in range(1, sheet.nrows):
            base.add_row(sheet.name, map(lambda x: x.value, rows[row_i]))

    base.execute('INSERT or IGNORE INTO "Товар" VALUES ("a1", "b", "c", "d", "e", "f")')
    base.cur.execute('SELECT * FROM "Магазин"')
    print(base.cur.fetchall())
    base.close()


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
