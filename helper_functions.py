import pandas as pd

# X_caregiver_patient_timeslot_room

SINGLE_TIME_SIZE_MINUTES = 15
START_TIME_HOUR = 8
END_TIME_HOUR = 17
TIME_ARRAY_SIZE = int((END_TIME_HOUR - START_TIME_HOUR) * 60.0 / SINGLE_TIME_SIZE_MINUTES)

CAREGIVERS_AMOUNT = 10
PATIENTS_AMOUNT = 11
ROOMS_AMOUNT = 12

ROOMS_ARRAY_SIZE = 1
TIMESLOTS_ARRAY_SIZE = ROOMS_AMOUNT
PATIENTS_ARRAY_SIZE = TIMESLOTS_ARRAY_SIZE * TIME_ARRAY_SIZE
CAREGIVERS_ARRAY_SIZE = PATIENTS_AMOUNT * ROOMS_AMOUNT * TIME_ARRAY_SIZE

def get_fixed_index(appointment_time, caregiver_i, patient_i, room_i):
    fixed_appointment_time = int((appointment_time - START_TIME_HOUR) * 60.0 / SINGLE_TIME_SIZE_MINUTES)
    total_index = caregiver_i * CAREGIVERS_ARRAY_SIZE + patient_i * PATIENTS_ARRAY_SIZE + fixed_appointment_time * TIMESLOTS_ARRAY_SIZE + room_i
    return total_index

def get_variables(fixed_index):
    caregiver_i = fixed_index // CAREGIVERS_ARRAY_SIZE
    fixed_index = fixed_index % CAREGIVERS_ARRAY_SIZE
    patient_i = fixed_index // PATIENTS_ARRAY_SIZE
    fixed_index = fixed_index % PATIENTS_ARRAY_SIZE
    fixed_appointment_time = fixed_index // TIMESLOTS_ARRAY_SIZE
    fixed_index = fixed_index % TIMESLOTS_ARRAY_SIZE
    room_i = fixed_index // ROOMS_ARRAY_SIZE

    appointment_time = START_TIME_HOUR + fixed_appointment_time * SINGLE_TIME_SIZE_MINUTES / 60.0
    return appointment_time, caregiver_i, patient_i, room_i

# Example usage
def test_example():
    appointment_time = 8.25
    caregiver_i = 2
    patient_i = 3
    room_i = 5

    fixed_index = get_fixed_index(appointment_time, caregiver_i, patient_i, room_i)
    print(get_variables(fixed_index))

def excel_sheets_to_items(excel_file="Rehabilitation Data.xlsx"):
    PATIENT_SHEET_NAME = "Patient Equipments"
    CAREGIVER_SHEET_NAME = "Caregiver Equipments"

    # Load the entire Excel file (all sheets)
    patients_sheet, caregivers_sheet  = pd.read_excel(excel_file, sheet_name=None).values()  # sheet_name=None loads all sheets
    
    equipment_list = patients_sheet.columns.tolist()[2:]
    patient_equipments_dict = {}
    caregiver_equipments_dict = {}
    patient_unavailable_times_dict = {}
    caregiver_unavailable_times_dict = {}
    
    for _, patient_row in patients_sheet.iterrows():
        patient_name = patient_row["Patient Name"]
        
        required_equipment = []
        for equipment in equipment_list:
            if not pd.isna(patient_row[equipment]):
                required_equipment.append((equipment, int(patient_row[equipment])))
        patient_equipments_dict[patient_name] = required_equipment

        if not pd.isna(patient_row["Unavailability Hours"]):
            unavailability_hours = []
            for current_unavailability_window in patient_row["Unavailability Hours"].split(", "):
                current_start_hour, current_end_hour = current_unavailability_window.split("-")
                current_unavailability_hours = [hour for hour in range(int(current_start_hour), int(current_end_hour))]
                unavailability_hours += current_unavailability_hours
            patient_unavailable_times_dict[patient_name] = unavailability_hours
        
    patients_list = list(patient_equipments_dict.keys())

    
    for _, caregiver_row in caregivers_sheet.iterrows():
        caregiver_name = caregiver_row["Caregiver Name"]
        
        treating_equipment = []
        if not pd.isna(caregiver_row["Treating Equipment"]):
            for equipment in caregiver_row["Treating Equipment"].split(", "):
                treating_equipment.append(equipment)
                if equipment not in equipment_list:
                    equipment_list.append(equipment)
        caregiver_equipments_dict[caregiver_name] = treating_equipment

        if not pd.isna(caregiver_row["Unavailability Hours"]):
            unavailability_hours = []
            for current_unavailability_window in caregiver_row["Unavailability Hours"].split(", "):
                current_start_hour, current_end_hour = current_unavailability_window.split("-")
                current_unavailability_hours = [hour for hour in range(int(current_start_hour), int(current_end_hour))]
                unavailability_hours += current_unavailability_hours
            caregiver_unavailable_times_dict[caregiver_name] = unavailability_hours

    caregivers_list = list(caregiver_equipments_dict.keys())

    return caregivers_list, patients_list, equipment_list, caregiver_equipments_dict, caregiver_unavailable_times_dict, patient_unavailable_times_dict, patient_equipments_dict



# print(result_dict)