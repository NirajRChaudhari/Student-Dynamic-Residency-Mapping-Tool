from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import random

def initialize():
    # Initialize global variables
    global students, timeSlots, organizations, orgNameToCode, timeCodeToSlot
    global dataExcelFile, processingWorkbook, studentPrefSheet, processingSheet

    # Dictionaries for storing data
    students = {}
    orgNameToCode = {}
    timeCodeToSlot = {}
    organizations = {}

    # Load the workbook containing student and organization data
    dataExcelFile = load_workbook("DataFiles/StudentResidencyData.xlsx")
    orgSheet = dataExcelFile["OrganizationDetails"]
    studentPrefSheet = dataExcelFile["StudentPreferenceDetails"]
    timeSlotSheet = dataExcelFile["TimeSlotDetails"]

    # Create a new workbook for processing the allocations
    processingWorkbook = Workbook()
    processingWorkbook.create_sheet("StudentPreferences", 0)
    processingSheet = processingWorkbook["StudentPreferences"]

    # Initialize organization name to code mapping
    for i in range(2, orgSheet.max_row + 1):
        orgName = orgSheet.cell(i, 1).value
        orgNameToCode[orgName] = {'code': orgSheet.cell(i, 2).value}

    # Initialize time code to slot mapping
    timeSlots = timeSlotSheet.max_row - 1
    for i in range(2, timeSlotSheet.max_row + 1):
        timeCode = timeSlotSheet.cell(i, 1).value
        timeCodeToSlot[timeCode] = {'slot': timeSlotSheet.cell(i, 2).value}

    # Initialize organizations with slots and students per slot
    for i in range(2, orgSheet.max_row + 1):
        orgCode = orgSheet.cell(i, 2).value
        organizations[orgCode] = {
            'name': orgSheet.cell(i, 1).value,
            'slotsAllocatedToOrg': orgSheet.cell(i, 3).value,
            'studentsPerSlot': orgSheet.cell(i, 4).value,  # Students per slot
            'allocatedStudents': 0,
            'studentsIDSlotMapping': {}
        }
        for j in range(1, organizations[orgCode]['slotsAllocatedToOrg']+1):
            organizations[orgCode]['studentsIDSlotMapping'][j] = []

    # Initialize students with preferences
    for i in range(2, studentPrefSheet.max_row + 1):
        uscId = studentPrefSheet.cell(i, 2).value
        students[uscId] = {
            'name': studentPrefSheet.cell(i, 1).value,
            'allocatedOrganizations': 0,
            'organizationsCodeSlotMapping': {},
            "preferences": {}
        }
        for slot in range(1, timeSlots+1):
            students[uscId]['organizationsCodeSlotMapping'][slot] = None

        prefNo = 1
        for j in range(3, studentPrefSheet.max_column + 1):
            orgName = studentPrefSheet.cell(i, j).value
            if orgName in orgNameToCode:
                students[uscId]["preferences"][prefNo] = orgNameToCode[orgName]['code']
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

    for i in range(2, studentPrefSheet.max_row + 1):
        processingSheet.cell(i, 1).value = studentPrefSheet.cell(i, 1).value
        processingSheet.cell(i, 2).value = studentPrefSheet.cell(i, 2).value
        for j in range(3, studentPrefSheet.max_column + 1):
            orgName = studentPrefSheet.cell(i, j).value
            if orgName in orgNameToCode:
                processingSheet.cell(i, j).value = orgNameToCode[orgName]['code']

def dynamic_allocation_of_students():
    # Dynamic allocation based on student preferences and organization capacity
    max_preferences = 10
    last_allocated_preference = {studentUSCId: 0 for studentUSCId in students}

    allocation_possible = True
    while allocation_possible:
        allocation_possible = False
        student_ids = list(students.keys())
        random.shuffle(student_ids)

        for studentUSCId in student_ids:
            if students[studentUSCId]['allocatedOrganizations'] >= timeSlots:
                continue  # Student already allocated maximum organizations

            for pref_index in range(last_allocated_preference[studentUSCId] + 1, max_preferences + 1):
                currentOrganization = students[studentUSCId]["preferences"].get(pref_index)
                if not currentOrganization or organizations.get(currentOrganization) is None:
                    continue  # Skip invalid preferences

                if organizations[currentOrganization]['allocatedStudents'] >= organizations[currentOrganization]['slotsAllocatedToOrg'] * organizations[currentOrganization]['studentsPerSlot']:
                    continue  # Organization at capacity

                for slot in range(1, timeSlots + 1):
                    if students[studentUSCId]['organizationsCodeSlotMapping'][slot] is not None:
                        continue  # Student already allocated to a slot

                    if len(organizations[currentOrganization]['studentsIDSlotMapping'][slot]) < organizations[currentOrganization]['studentsPerSlot']:
                        # Allocate student to this organization and slot
                        organizations[currentOrganization]['allocatedStudents'] += 1
                        students[studentUSCId]['allocatedOrganizations'] += 1
                        organizations[currentOrganization]['studentsIDSlotMapping'][slot].append(studentUSCId)
                        students[studentUSCId]['organizationsCodeSlotMapping'][slot] = currentOrganization
                        
                        last_allocated_preference[studentUSCId] = pref_index
                        allocation_possible = True  # Allocation was successful
                        break

def populate_processing_workbook():
    # Populate the processing workbook with allocation results
    processingWorkbook.create_sheet("StudentsMapping", 1)
    studentsMappingSheet = processingWorkbook["StudentsMapping"]
    studentsMappingSheet.cell(1, 1).value = "USC ID"
    studentsMappingSheet.cell(1, 2).value = "Student Name"
    studentsMappingSheet.cell(1, 1).font = Font(color="00FF0000", bold=True)
    studentsMappingSheet.cell(1, 2).font = Font(color="00FF0000", bold=True)

    for i in range(1, timeSlots+1):
        studentsMappingSheet.cell(1, i + 2).value = timeCodeToSlot[i]['slot']
        studentsMappingSheet.cell(1, i + 2).font = Font(color="00FF0000", bold=True)

    for i, (uscId, student) in enumerate(students.items(), start=2):
        studentsMappingSheet.cell(i, 1).value = uscId
        studentsMappingSheet.cell(i, 2).value = student['name']

        for slot in range(1, timeSlots+1):
            if student['organizationsCodeSlotMapping'][slot]:
                orgCode = student['organizationsCodeSlotMapping'][slot]
                studentsMappingSheet.cell(i, slot + 2).value = organizations[orgCode]['name']
            else:
                studentsMappingSheet.cell(i, slot + 2).value = ""

    processingWorkbook.create_sheet("OrganizationMapping", 2)
    organizationMappingSheet = processingWorkbook["OrganizationMapping"]
    organizationMappingSheet.cell(1, 1).value = "Organization Code"
    organizationMappingSheet.cell(1, 2).value = "Organization Name"
    organizationMappingSheet.cell(1, 1).font = Font(color="00FF0000", bold=True)
    organizationMappingSheet.cell(1, 2).font = Font(color="00FF0000", bold=True)

    for i in range(1, timeSlots+1):
        organizationMappingSheet.cell(1, i + 2).value = timeCodeToSlot[i]['slot']
        organizationMappingSheet.cell(1, i + 2).font = Font(color="00FF0000", bold=True)

    for i, (code, org) in enumerate(organizations.items(), start=2):
        organizationMappingSheet.cell(i, 1).value = code
        organizationMappingSheet.cell(i, 2).value = org['name']

        for slot in range(1, org['slotsAllocatedToOrg']+1):
            if slot in org['studentsIDSlotMapping'] and org['studentsIDSlotMapping'][slot]:
                studentNames = [students[studentID]['name'] for studentID in org['studentsIDSlotMapping'][slot]]
                organizationMappingSheet.cell(i, slot + 2).value = ", ".join(studentNames)
            else:
                organizationMappingSheet.cell(i, slot + 2).value = ""

def main():
    initialize()
    preprocessing_student_preferences_sheet()
    dynamic_allocation_of_students()
    populate_processing_workbook()
    processingWorkbook.save("DataFiles/ProcessingWorkbook_1.xlsx")

if __name__ == "__main__":
    main()
