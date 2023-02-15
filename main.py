from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


def initialize():
    global students
    global timeSlots
    global organizations
    global orgNameToCode
    global timeCodeToSlot

    global dataExcelFile
    global processingWorkbook
    global studentPrefSheet
    global processingSheet

    students = {}
    orgNameToCode = {}
    timeCodeToSlot = {}
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

    # Initialize timeCodeToSlot dictionary
    timeSlots = timeSlotSheet.max_row - 1
    for i in range(2, timeSlotSheet.max_row + 1):
        timeCode = timeSlotSheet.cell(i, 2).value
        timeCodeToSlot[timeCode] = {
            'code': timeSlotSheet.cell(i, 1).value
        }

    # Initialize organizations dictionary
    for i in range(2, orgSheet.max_row + 1):
        orgCode = orgSheet.cell(i, 2).value

        organizations[orgCode] = {}
        organizations[orgCode]['name'] = orgSheet.cell(i, 1).value
        organizations[orgCode]['allocatedStudents'] = 0
        organizations[orgCode]['studentsIDSlotMapping'] = {}
        for i in range(0, timeSlots):
            organizations[orgCode]['studentsIDSlotMapping'][i] = None

    # Initialize students dictionary
    for i in range(2, studentPrefSheet.max_row + 1):
        uscId = studentPrefSheet.cell(i, 1).value

        students[uscId] = {}
        students[uscId]['name'] = studentPrefSheet.cell(i, 2).value
        students[uscId]['allocatedOrganizations'] = 0
        students[uscId]['organizationsCodeSlotMapping'] = {}
        for slot in range(0, timeSlots):
            students[uscId]['organizationsCodeSlotMapping'][slot] = None

        prefNo = 1

        students[uscId]["preferences"] = {}
        for j in range(3, studentPrefSheet.max_column + 1):
            students[uscId]["preferences"][prefNo] = orgNameToCode[studentPrefSheet.cell(
                i, j).value]['code']
            prefNo += 1
        print("\n\n")


def preprocessing_student_preferences_sheet():
    global students
    global orgNameToCode
    global processingSheet
    global studentPrefSheet

    # Print the dictionary orgNameToCode with proper formatting
    print("\n\n\nOrganizations Name to Code:\n")
    for key, value in orgNameToCode.items():
        print(key, value['code'])

    # Print the dictionary timeCodeToSlot with proper formatting
    print("\n\n\nTime Slots:\n")
    for key, value in timeCodeToSlot.items():
        print(key, value['code'])

    # Print the dictionary students with proper formatting
    print("\n\n\nStudents:\n")
    for key, value in students.items():
        print(key, value['name'])
        for prefNo, pref in value.items():
            if prefNo != 'name':
                print(prefNo, pref)
        print("\n")

    # Print the dictionary organizations with proper formatting
    print("\n\nOrganizations:\n")
    for key, value in organizations.items():
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


def dynamic_allocation_of_students():
    global students
    global organizations

    for index in range(1, studentPrefSheet.max_column - 1):
        for studentUSCId in students:

            print(studentUSCId)
            if (students[studentUSCId]["preferences"].get(index)):
                currentOrganization = students[studentUSCId]["preferences"].get(
                    index)
                if (organizations.get(currentOrganization) == None):
                    print(
                        "\nWRONG ORGANIZATION NAME AT PREFERENCE "+str(1)+" FOR STUDENT WITH USC ID : "+str(studentUSCId)+"\n")
                    return

            if (organizations.get(students[studentUSCId]["preferences"].get(index))['allocatedStudents'] == timeSlots):
                print(
                    "\nORGANIZATION "+currentOrganization+" HAS ALL TIME SLOTS FULL SO PREFERENCE : "+str(index)+" OF STUDENT WITH USC ID : "+str(studentUSCId)+" CANNOT BE CONSIDERED\n")
                continue

            if (students[studentUSCId]['allocatedOrganizations'] == timeSlots):
                print(
                    "\nSTUDENT WITH USC ID : "+str(studentUSCId)+" HAS ALL TIME SLOTS FULL SO SKIP CURRENT PREFERENCE\n")

            orgAssigned = False
            for slot in range(0, timeSlots):
                if (students[studentUSCId]['organizationsCodeSlotMapping'][slot] == None):
                    if (organizations[currentOrganization]['studentsIDSlotMapping'][slot] == None):
                        orgAssigned = True

                        organizations[currentOrganization]['allocatedStudents'] += 1
                        students[studentUSCId]['allocatedOrganizations'] += 1

                        organizations[currentOrganization]['studentsIDSlotMapping'][slot] = studentUSCId

                        students[studentUSCId]['organizationsCodeSlotMapping'][slot] = currentOrganization

                        break

            if (not orgAssigned):
                print("\nNO COMPATIBILITY IN STUDENT WITH USC ID : "+str(studentUSCId) +
                      " WITH PREFERENCE "+str(index)+" AND ORGANIZATION "+currentOrganization+" FOUND\n")


def populate_processing_workbook():
    global processingWorkbook

    processingWorkbook.create_sheet("StudentsMapping", 1)
    studentsMappingSheet = processingWorkbook["StudentsMapping"]

    # Write the headings to the students mapping sheet
    studentsMappingSheet.cell(1, 1).value = "USC ID"
    studentsMappingSheet.cell(1, 2).value = "Student Name"
    for i in range(0, timeSlots):
        studentsMappingSheet.cell(
            1, i + 3).value = timeCodeToSlot[i]['code']

    # Write the students dictionary to the students mapping sheet
    for i, (key, value) in enumerate(students.items()):
        studentsMappingSheet.cell(i + 2, 1).value = key
        studentsMappingSheet.cell(i + 2, 2).value = value['name']
        for slot, org in value['organizationsCodeSlotMapping'].items():
            studentsMappingSheet.cell(
                i + 2, slot + 3).value = organizations[org]['name']

    processingWorkbook.create_sheet("OrganizationMapping", 2)
    organizationMappingSheet = processingWorkbook["OrganizationMapping"]

    # Write the organizations dictionary to the organization mapping sheet
    organizationMappingSheet.cell(1, 1).value = "Organization Code"
    organizationMappingSheet.cell(1, 2).value = "Organization Name"

    for i in range(0, timeSlots):
        organizationMappingSheet.cell(
            1, i + 3).value = timeCodeToSlot[i]['code']

    for i, (key, value) in enumerate(organizations.items()):
        organizationMappingSheet.cell(i + 2, 1).value = key
        organizationMappingSheet.cell(i + 2, 2).value = value['name']
        for slot, student in value['studentsIDSlotMapping'].items():
            organizationMappingSheet.cell(i + 2, slot + 3).value = student


def main():
    global processingWorkbook
    initialize()
    dynamic_allocation_of_students()
    preprocessing_student_preferences_sheet()

    populate_processing_workbook()
    # Save the processing sheet
    processingWorkbook.save("DataFiles/ProcessingWorkbook.xlsx")


if __name__ == "__main__":
    print("\n")
    main()
    print("\n\n")
