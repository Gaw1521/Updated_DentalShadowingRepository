# Dental Shadowing List

# Update Date: 01/04/2022


import xlsxwriter


#  Create file (workbook) and worksheet
DentalWorkbook = xlsxwriter.Workbook("DentalShadowingLog_InPerson.xlsx")
DentalWorksheet = DentalWorkbook.add_worksheet()


#  Declare data
Dr = ["Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Shaun Small, DMD", "Dr. Gerald G. White, DDS",
      "Dr. Gerald G. White, DDS", "Dr. Shaun Small, DMD", "Dr. Shaun Small, DMD",
      "Dr. Shaun Small, DMD", "Dr. Gerald G. White, DMD", "Dr. Shaun Small, DMD",
      "Dr. Gerald G. White", "Dr. Gerald G. White", "Dr. Gerald G. White",
      "Dr. Shaun Small, DMD", "Dr. Gerald G. White", "Dr. Shaun Small, DMD",
      "Dr. Shaun Small, DMD", "Dr. Shaun Small, DMD",  "Dr. Shaun Small, DMD",
      "Dr. Shaun Small, DMD", "Dr. Shaun Small, DMD", "Dr. Gerald G. White",
      "Dr. Gerald G. White", "Dr. Forrest Crabtree, DMD", "Dr. Forrest Crabtree, DMD",
      "Dr. Forrest Crabtree, DMD", "Dr. Forrest Crabtree, DMD", "Dr. Forrest Crabtree, DMD",


      ]


GD_Specialty = ["General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist",  "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist",  "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",
                "General Dentist", "General Dentist", "General Dentist",

                ]

Date = ["08-31-2020",
        "09-01-2020",
        "09-02-2020",
        "09-03-2020",
        "09-08-2020",
        "09-10-2020",
        "09-28-2020",
        "09-30-2020",
        "10-12-2020",
        "10-13-2020",
        "10-25-2020",
        "10-26-2020",
        "10-27-2020",
        "10-28-2020",
        "11-19-2020",
        "11-29-2020",
        "12-07-2020",
        "12-08-2020",
        "12-09-2020",
        "12-14-2020",
        "12-13-2020",
        "12-15-2020",
        "12-16-2020",
        "12-21-2020",
        "12-23-2020",
        "01-03-2021",
        "01-06-2021",
        "01-07-2021",
        "01-12-2021",
        "01-13-2021",
        "01-14-2021",
        "01-18-2021",
        "01-28-2021",
        "02-02-2021",
        "02-24-2021",
        "03-04-2021",
        "03-18-2021",
        "03-25-2021",
        "05-17-2021",
        "05-19-2021",
        "12-20-2021",
        "12-23-2021",
        "12-27-2021",
        "01-06-2022",
        "01-07-2022",

        ]

Hours = ["09:30",
         "09:30",
         "09:30",
         "09:30",
         "09:30",
         "09:30",
         "09:30",
         "09:30",
         "07:30",
         "04:00",
         "04:00",
         "04:00",
         "04:00",
         "05:00",
         "05:00",
         "04:00",
         "04:00",
         "04:00",
         "05:00",
         "04:30",
         "04:30",
         "02:00",
         "04:30",
         "04:30",
         "04:30",
         "03:00",
         "04:30",
         "04:00",
         "05:00",
         "03:00",
         "04:30",
         "07:00",
         "03:00",
         "03:00",
         "03:00",
         "3:00",
         "3:00",
         "3:00",
         "3:00",
         "3:00",
         "5:00",
         "5:00",
         "5:00",
         "5:00",
         "5:00",


         ]


#  Column Headers
DentalWorksheet.write("A1", "Dr.")
DentalWorksheet.write("B1", "GD/Specialty")
DentalWorksheet.write("C1", "Date")
DentalWorksheet.write("D1", "Hours")


#  Add a bold format to use to highlight cells.
bold = DentalWorkbook.add_format({'bold': True})

DentalWorksheet.write('A1', 'Dr', bold)
DentalWorksheet.write('B1', 'GD/Specialty', bold)
DentalWorksheet.write('C1', 'Date', bold)
DentalWorksheet.write('D1', 'Hours', bold)

#  Setting Color to a Cell
        #  cell_format = DentalWorkbook.add_format()

        #  cell_format.set_pattern(1)
        #  cell_format.set_bg_color('orange')

        #  DentalWorksheet.write('A1', 'Dr', cell_format, bold)

#  Write data to file
for item in range(len(Dr)):
    DentalWorksheet.write(item+1, 0, Dr[item])
    DentalWorksheet.write(item + 1, 1, GD_Specialty[item])
    DentalWorksheet.write(item + 1, 2, Date[item])
    DentalWorksheet.write(item + 1, 3, Hours[item])


#  Write data to file
#  EmailWorksheet.write(1, 0, email[0])
#  EmailWorksheet.write(2, 0, email[1])

#  EmailWorksheet.write(1, 1, password[0])
#  EmailWorksheet.write(2, 1, password[1])

DentalWorkbook.close()
