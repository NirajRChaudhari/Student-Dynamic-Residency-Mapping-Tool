from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


def initialize():
    global students
    global timeSlots
    global organizations
    global orgNameToCode
    global timeSlotToCode

    global dataExcelFile
    global processingWorkbook

    global studentPrefSheet
    global processingSheet

    students = {}
    orgNameToCode = {}
    timeSlotToCode = {}
    organizations = {}
    dataExcelFile = None
    processingWorkbook = None

    dataExcelFile = load_workbook("DataFiles/StudentResidencyData.xlsx")
    orgSheet = dataExcelFile["OrganizationDetails"]
    studentPrefSheet = dataExcelFile["StudentPreferenceDetails"]
    timeSlotSheet = dataExcelFile["TimeSlotDetails"]

    processingWorkbook = Workbook()
    processingWorkbook.create_sheet("StudentPreferences", 0)
    processingSheet = processingWorkbook["StudentPreferences"]

    # Initialize orgNameToCode dictionary
    for i in range(2, orgSheet.max_row + 1):
        orgName = orgSheet.cell(i, 1).value
        orgNameToCode[orgName] = {
            'code': orgSheet.cell(i, 2).value
        }

    # Initialize timeSlotToCode dictionary
    timeSlots = timeSlotSheet.max_row - 1
    for i in range(2, timeSlotSheet.max_row + 1):
        timeSlot = timeSlotSheet.cell(i, 1).value
        timeSlotToCode[timeSlot] = {
            'code': timeSlotSheet.cell(i, 2).value
        }

    # Initialize organizations dictionary
    for i in range(2, orgSheet.max_row + 1):
        orgCode = orgSheet.cell(i, 2).value

        organizations[orgCode] = {}
        organizations[orgCode]['name'] = orgSheet.cell(i, 1).value
        organizations[orgCode]['allocatedStudents'] = 0
        organizations[orgCode]['students'] = []

    # Initialize students dictionary
    for i in range(2, studentPrefSheet.max_row + 1):
        uscId = studentPrefSheet.cell(i, 1).value

        students[uscId] = {}
        students[uscId]['name'] = studentPrefSheet.cell(i, 2).value
        students[uscId]['allocatedOrganizations'] = 0
        students[uscId]['organizations'] = []
        prefNo = 1

        for j in range(3, studentPrefSheet.max_column + 1):
            students[uscId][prefNo] = orgNameToCode[studentPrefSheet.cell(
                i, j).value]['code']
            prefNo += 1


def preprocessing_student_preferences_sheet():
    global students
    global orgNameToCode
    global processingSheet
    global studentPrefSheet

    # Print the dictionary orgNameToCode with proper formatting
    print("\n\n\nOrganizations Name to Code:\n")
    for key, value in orgNameToCode.items():
        print(key, value['code'])

    # Print the dictionary timeSlotToCode with proper formatting
    print("\n\n\nTime Slots:\n")
    for key, value in timeSlotToCode.items():
        print(key, value['code'])

    # Print the dictionary students with proper formatting
    print("\n\n\nStudents:\n")
    for key, value in students.items():
        print(key, value['name'])
        for prefNo, pref in value.items():
            if prefNo != 'name':
                print(prefNo, pref)
        print("\n")

    # Copy the student preferences headings to the processing sheet
    for i in range(1, studentPrefSheet.max_column + 1):
        processingSheet.cell(1, i).value = studentPrefSheet.cell(1, i).value
        processingSheet.cell(1, i).font = Font(bold=True)

    # Copy the student preferences to the processing sheet
    for i in range(2, studentPrefSheet.max_row + 1):

        processingSheet.cell(i, 1).value = studentPrefSheet.cell(i, 1).value
        processingSheet.cell(i, 2).value = studentPrefSheet.cell(i, 2).value

        for j in range(3, studentPrefSheet.max_column + 1):
            processingSheet.cell(
                i, j).value = orgNameToCode[studentPrefSheet.cell(i, j).value]['code']


def main():
    global processingWorkbook
    initialize()
    preprocessing_student_preferences_sheet()

    # Save the processing sheet
    processingWorkbook.save("DataFiles/ProcessingWorkbook.xlsx")


if __name__ == "__main__":
    print("\n")
    main()
    print("\n\n")
