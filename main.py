import openpyxl as xl
from tkinter import filedialog
import re
from datetime import datetime as dt
import datetime
import random


class TimeTable:
    def __init__(
        self, sheet: xl.worksheet.worksheet.Worksheet, tt_range: str, fix_rows=True
    ):
        self.rows = sheet[tt_range]
        self.fixrows(fix_rows)

    def fixrows(self, fix_rows):
        """Structure rows in the proper format for further processing)"""
        if fix_rows:
            rr = []
            for row in self.rows[1:]:
                rr.append(row[1:])

            self.rows = rr

        # Little more structuring.
        # Multiple possible classes in the same slot.
        self.rows = [[cell.value for cell in row] for row in self.rows]
        self.rows = [
            [cell.rstrip("\n").split("\n") if cell else [] for cell in row]
            for row in self.rows
        ]

    def removeOtherCourses(self):
        """Removes other courses you are not registerd to
        Ensure you have correctly added your courses to your_courses.txt
        """

        input(
            "Ensure you have added your course codes to the file 'your_courses.txt' before continuing. Press Enter to continue."
        )

        with open("your_courses.txt", "r") as file:
            tlist = file.readlines()

        clist = []
        for c in tlist:
            if c[0] != "#" and c != "\n":
                clist.append(c.rstrip("\n").strip())

        self.clist = clist
        # print(self.clist)

        def isValidCourse(ele):
            """Inner Function verifies if course valid"""
            return any(map(lambda c: c in ele, self.clist))

        self.rows_updated = []
        for row in self.rows:
            nrow = []
            for cell in row:
                ncell = list(filter(isValidCourse, cell))
                nrow.append(ncell)
            self.rows_updated.append(nrow)

    def getCourseDetails(self):
        """Structures courses to have fields for .ics files."""
        timings = [
            "8:30-9:25",
            "9:30-10:25",
            "10:40-11:35",
            "11:40-12:35",
            "12:40-1:30",
            "13:30-14:25",
            "14:30-15:25",
            "15:40-16:35",
            "16:40-17:35",
            "17:40-18:35",
        ]
        timings = [timing.split("-") for timing in timings]

        days = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
        ]  # Not used.
        dates = [
            "15/01/2024",
            "16/01/2024",
            "17/01/2024",
            "18/01/2024",
            "19/01/2024",
            "20/01/2024",
        ]  # Calendar starts from Monday 15th Jan 2024

        merged_classes = []

        def getCourse(row, col):
            if row in merged_classes:
                return

            class_ele = "\n".join(
                self.rows_updated[i][j]
            )  # Incase there are multiple courses (cuz Ben10 👍)

            if not class_ele:
                return

            def parse_room_and_class_element(string):
                room = re.search("{(.*?)}", string)[0]

                class_ele = string.replace(room, "")
                class_ele = string.strip(": ")

                return (room, class_ele)

            room, class_ele = parse_room_and_class_element(class_ele)

            room = room[1:-1]

            startt, endt = timings[row]
            date = dates[col]

            startt = dt.strptime(f"{date} {startt}", "%d/%m/%Y %H:%M")

            next_class_ele = "\n".join(self.rows_updated[i + 1][j])
            merged_endtime = endt

            if next_class_ele:
                _, next_class_desc = parse_room_and_class_element(next_class_ele)

                if class_ele == next_class_desc:
                    _, next_endt = timings[row + 1]
                    merged_endtime = next_endt
                    merged_classes.append(row + 1)

            endt = dt.strptime(f"{date} {merged_endtime}", "%d/%m/%Y %H:%M")

            for cc in self.clist:
                if cc in class_ele:
                    course_code = cc

            return startt, endt, course_code, class_ele, room

        self.course_dets = []
        si, sj = len(self.rows_updated), len(self.rows_updated[0])
        for i in range(si):
            for j in range(sj):
                course = getCourse(i, j)
                if course:
                    self.course_dets.append(list(course))

    def rewriteCourseNames(self):
        print("Rewriting course codes as their course names...")
        with open("course_codes.txt", "r") as file:
            cc = file.readlines()

        course_codes = []
        for c in cc:
            if c[0] != "#" and c != "\n":
                course_codes.append(c.rstrip("\n").split(": "))

        course_codes = dict((cc, cname) for cc, cname in course_codes)

        self.course_codes = course_codes

        # Replace course code with course name
        for i in range(len(self.course_dets)):
            dets = self.course_dets[i]
            cc = dets[2]
            self.course_dets[i][3] = dets[3].replace(cc, course_codes[cc])

    def makeICS(self):
        ics_header = """BEGIN:VCALENDAR
VERSION:2.0
PRODID:- GitHub: BaronTime-Buggermenot, Time Table
X-WR-CALNAME;VALUE=TEXT:Generated TimeTable Baron-Buggermenot [GitHub]

BEGIN:VTIMEZONE
TZID:Asia/Calcutta
X-LIC-LOCATION:Asia/Calcutta
BEGIN:STANDARD
TZOFFSETFROM:+0530
TZOFFSETTO:+0530
TZNAME:IST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
"""

        ics_events = "\n"
        random.seed(170008200022)

        for week in range(21):  # Weeks until first week of June
            for startt, endt, course_code, class_desc, room in self.course_dets:
                startt = startt + datetime.timedelta(weeks=week)
                endt = endt + datetime.timedelta(weeks=week)
                day = dt.strftime(startt, "%w")

                ics_events += "BEGIN:VEVENT\n"
                ics_events += f"SUMMARY:{class_desc}\n"
                ics_events += f"DTSTAMP:20002208T000000\n"
                ics_events += f"DESCRIPTION:Class Time!\n"
                ics_events += f"DTSTART:{dt.strftime(startt, '%Y%m%dT%H%M%S')}\n"
                ics_events += f"DTEND:{dt.strftime(endt, '%Y%m%dT%H%M%S')}\n"
                ics_events += f"LOCATION:{room}\n"
                ics_events += f"UID:{str(week).zfill(2)}{str(random.randint(0, 100)).zfill(2)}{day}{course_code}\n"
                ics_events += "END:VEVENT\n\n"

        with open("TimeTable Baron.ics", "w") as timeTable:
            timeTable.writelines(ics_header + ics_events + "END:VCALENDAR")

    def compile(self):
        self.removeOtherCourses()  # Removes all other courses not in your_courses.txt
        self.getCourseDetails()  # Structures and returns course details as: start_time, end_time, course_code, description, room
        self.rewriteCourseNames()  # Replaces course_code with course name in decription
        self.makeICS()  # Generates .ics file for the entire semester.


if __name__ == "__main__":
    xl_file = filedialog.askopenfilename()
    print("Selected File:", xl_file)

    xl_file = xl.load_workbook(xl_file)

    # Time Table Range and verification
    tt_range = "A4:G14"
    if input(f"TimeTable Range: {tt_range} Correct?"):
        tt_range = input("Input new Range: ")
        print("Range Update to:", tt_range)

    TT = TimeTable(xl_file.active, tt_range, fix_rows=True)
    TT.compile()
