from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Color


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
        timeCode = timeSlotSheet.cell(i, 1).value
        timeCodeToSlot[timeCode] = {
            'slot': timeSlotSheet.cell(i, 2).value
        }

    # Initialize organizations dictionary
    for i in range(2, orgSheet.max_row + 1):
        orgCode = orgSheet.cell(i, 2).value

        organizations[orgCode] = {}
        organizations[orgCode]['name'] = orgSheet.cell(i, 1).value
        organizations[orgCode]['slotsAllocatedToOrg'] = orgSheet.cell(
            i, 3).value
        organizations[orgCode]['allocatedStudents'] = 0
        organizations[orgCode]['studentsIDSlotMapping'] = {}

        for i in range(1, organizations[orgCode]['slotsAllocatedToOrg']+1):
            organizations[orgCode]['studentsIDSlotMapping'][i] = None

    # Initialize students dictionary
    for i in range(2, studentPrefSheet.max_row + 1):
        uscId = studentPrefSheet.cell(i, 1).value

        students[uscId] = {}
        students[uscId]['name'] = studentPrefSheet.cell(i, 2).value
        students[uscId]['allocatedOrganizations'] = 0
        students[uscId]['organizationsCodeSlotMapping'] = {}
        for slot in range(1, timeSlots+1):
            students[uscId]['organizationsCodeSlotMapping'][slot] = None

        students[uscId]["preferences"] = {}

        prefNo = 1
        for j in range(3, studentPrefSheet.max_column + 1):
            students[uscId]["preferences"][prefNo] = orgNameToCode[studentPrefSheet.cell(
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

    # Print the dictionary timeCodeToSlot with proper formatting
    print("\n\n\nTime Slots:\n")
    for key, value in timeCodeToSlot.items():
        print(key, value['slot'])

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

            # Check if the student has any more preferences and Fetch current organization preference
            currentOrganization = students[studentUSCId]["preferences"].get(
                index)

            if (currentOrganization != None and currentOrganization != "" and currentOrganization != "None"):
                if (organizations.get(currentOrganization) == None):
                    print(
                        "\nWRONG ORGANIZATION NAME AT PREFERENCE "+str(1)+" FOR STUDENT WITH USC ID : "+str(studentUSCId)+"\n")
                    return
            else:
                print(
                    "\nSTUDENT WITH USC ID : "+str(studentUSCId)+" HAS NO MORE PREFERENCES\n")
                continue

            if (organizations.get(currentOrganization)['allocatedStudents'] == organizations.get(currentOrganization)['slotsAllocatedToOrg']):
                print(
                    "\nORGANIZATION "+currentOrganization+" HAS IT'S ALL TIME SLOTS FULL SO PREFERENCE : "+str(index)+" OF STUDENT WITH USC ID : "+str(studentUSCId)+" CANNOT BE CONSIDERED\n")
                continue

            if (students[studentUSCId]['allocatedOrganizations'] == timeSlots):
                print(
                    "\nSTUDENT WITH USC ID : "+str(studentUSCId)+" HAS ALL TIME SLOTS FULL SO SKIP CURRENT PREFERENCE\n")

            orgAssigned = False
            for slot in range(1, timeSlots+1):

                if (slot > organizations[currentOrganization]['slotsAllocatedToOrg']):
                    break

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

    studentsMappingSheet.cell(1, 1).font = Font(color="00FF0000", bold=True)
    studentsMappingSheet.cell(1, 2).font = Font(color="00FF0000", bold=True)

    for i in range(1, timeSlots+1):
        studentsMappingSheet.cell(
            1, i + 2).value = timeCodeToSlot[i]['slot']

        studentsMappingSheet.cell(
            1, i + 2).font = Font(color="00FF0000", bold=True)

    # Write the students dictionary to the students mapping sheet
    for i, (uscId, student) in enumerate(students.items()):
        studentsMappingSheet.cell(i + 2, 1).value = uscId
        studentsMappingSheet.cell(i + 2, 2).value = student['name']

        for slot in range(1, timeSlots+1):
            if (student['organizationsCodeSlotMapping'][slot] == None):
                # No organization assigned
                studentsMappingSheet.cell(i + 2, slot + 2).value = ""
            else:
                studentsMappingSheet.cell(
                    i + 2, slot + 2).value = organizations[student['organizationsCodeSlotMapping'][slot]]['name']

    processingWorkbook.create_sheet("OrganizationMapping", 2)
    organizationMappingSheet = processingWorkbook["OrganizationMapping"]

    # Write the organizations dictionary to the organization mapping sheet
    organizationMappingSheet.cell(1, 1).value = "Organization Code"
    organizationMappingSheet.cell(1, 2).value = "Organization Name"

    organizationMappingSheet.cell(1, 1).font = Font(
        color="00FF0000", bold=True)
    organizationMappingSheet.cell(1, 2).font = Font(
        color="00FF0000", bold=True)

    for i in range(1, timeSlots+1):
        organizationMappingSheet.cell(
            1, i + 2).value = timeCodeToSlot[i]['slot']
        organizationMappingSheet.cell(
            1, i + 2).font = Font(color="00FF0000", bold=True)

    for i, (code, org) in enumerate(organizations.items()):
        organizationMappingSheet.cell(i + 2, 1).value = code
        organizationMappingSheet.cell(i + 2, 2).value = org['name']

        for slot in range(1, timeSlots+1):
            if (slot > org['slotsAllocatedToOrg']):
                organizationMappingSheet.cell(
                    i + 2, slot + 2).value = "NOT_AVAILABLE"
            elif (org['studentsIDSlotMapping'][slot] == None):
                organizationMappingSheet.cell(
                    i + 2, slot + 2).value = ""  # No student assigned
            else:
                organizationMappingSheet.cell(
                    i + 2, slot + 2).value = students[org['studentsIDSlotMapping'][slot]]['name']


def main():
    global processingWorkbook
    initialize()
    dynamic_allocation_of_students()
    preprocessing_student_preferences_sheet()

    populate_processing_workbook()
    # Save the processing sheet
    processingWorkbook.save("DataFiles/ProcessingWorkbook_1.xlsx")


if __name__ == "__main__":
    print("\n")
    main()
    print("\n\n")
