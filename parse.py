import csv
from email.policy import default
from traceback import print_exc
from openpyxl import Workbook
import openpyxl.styles.colors as xlcolour
from openpyxl.styles import PatternFill, Alignment


HEIGHT      = 4 # events start at 4 different times in hour
DAY_WIDTH   = 6
FILE_NAME   = "out1.xlsx"

ROW_BASE = 8*4 - 4

class Event:
    def __init__(self, row: list[str]) -> None:
        (self.name,
        self.group,
        self.type,
        self.title,
        self.info,
        self.wday,
        self.first_day,
        self.last_day,
        self.start_time,
        self.end_time,
        self.place,
        self.capacity,
        self.teacher,
        self.email,
        self.required_services,
        self.accepted,
        self.artefact) = row
    
    def cell_column(self) -> tuple[int, int]:
        """
        Generate column range for excel
        return (
            start_column,
            end_column
        )
        """
        # First generate width  -> if LECTURE   width 6
        #                       -> if CWA       width 2
        #                       -> if CWL       width 1

        width = {
            "Wykład": 6,
            "CWA":  3,
            "CWL":  1
        }.get(self.type, 1)

        cell = {
            "Pn": 0,
            "Wt": 1,
            "Śr": 2,
            "Cz": 3,
            "Pt": 4
        }.get(self.wday, None) * DAY_WIDTH + 1 + (int(self.group.strip("a")) - 1) * width

        return int(cell), int(cell + width - 1)

    def cell_row(self) -> int:
        """Generate row range for excel
        return (
            row_start,
            row_end
        )"""
        h, m = self.start_time.split(":")
        h = int(h)
        m = int(m)

        start = h * 4 + m / 15 - ROW_BASE

        h, m = self.end_time.split(":")
        h = int(h)
        m = int(m)

        end = h * 4 + m / 15 - ROW_BASE

        return int(start), int(end) - 1

    def colour(self):
        return {
            "Wykład": "FF3333",
            "CWA": "33FF33",
            "CWL": "3333FF"
        }.get(self.type, "333333")

    def value(self) -> str:
        return ", ".join([ self.type, self.title, self.teacher, self.start_time, self.end_time])

with open('events.csv', 'r', encoding='utf-8') as f:
    events = csv.reader(f)

    for row in events:
        for i, x in enumerate(row):
            print(i, x)
        break

    xl = Workbook()
    sh = xl.active

    for i, event in enumerate(events):
        if i == 0:
            print(event)
            continue
        try:
            e = Event(event)
            if e.type == "Lektorat":
                print("continue lektorat:", event)
                continue

            print("happeinng")
            start_col, end_col = e.cell_column()
            start_row, end_row = e.cell_row()


            sh.merge_cells(start_row=start_row, end_row=end_row, start_column=start_col, end_column=end_col)

            c = sh.cell(row=start_row, column=start_col)
            print(c)

            c.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
            c.fill = PatternFill(patternType='solid',  fgColor=xlcolour.Color(rgb="00"+e.colour()))
            c.value = e.value()
            print("Success: ", i)

        except:
            print_exc()
            print(i, event)
            xl.save(FILE_NAME)
            exit()

    xl.save(FILE_NAME)



        




