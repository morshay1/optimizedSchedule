import pandas as pd
import re
from helper_functions import excel_sheets_to_items
from create_schedule import save_schedule_to_excel

T = range(8, 18)

def replace_unavailable_patient_with_matching_available_one(input_file_path, schedule, patients_unavailability_list, output_file_path):
    # Define sets
    C, P, E, Ce, caregiver_unavailability, patient_unavailability, patient_equipment_mapping = \
        excel_sheets_to_items(input_file_path)

    # Read the Unscheduled Patients sheet directly from the Excel file
    unscheduled_patients_df = pd.read_excel(schedule, sheet_name="Unscheduled Patients")

    # Convert the data to a dictionary for easier processing
    patients_unscheduled = {}
    for _, row in unscheduled_patients_df.iterrows():
        patient = row["Patient"]
        equipment = row["Equipment"]
        count = row["Count"]
        if patient not in patients_unscheduled:
            patients_unscheduled[patient] = []
        patients_unscheduled[patient].append((equipment, count))

    schedule_file_path = create_schedule_from_excel(schedule)
    updated_schedule = reformat_schedule_with_regex(schedule_file_path)

    caregivers_list = find_caregivers_for_patients(updated_schedule, patients_unavailability_list)

    # Iterate over the caregivers and find matching patients from unscheduled list
    for caregiver, unavailable_time, unavailable_patient in caregivers_list:
        available_equipment = Ce.get(caregiver, [])  # Get equipment caregiver is qualified to use

        # Find a patient from patients_unscheduled who uses the matching equipment
        replacement_found = False
        for unscheduled_patient, equipment_list in patients_unscheduled.items():
            for eq, _ in equipment_list:
                if eq in available_equipment:
                    # Replace the patient in the schedule
                    print(
                        f"Replacing {unavailable_patient} with {unscheduled_patient} at {unavailable_time} using equipment {eq}")

                    # Update schedule by replacing patient_info
                    caregivers = schedule[unavailable_time]
                    for caregiver_key, patient_info in caregivers.items():
                        if caregiver_key == caregiver:
                            new_patient_info = re.sub(f"{unavailable_patient}, Equipment\d+",
                                                      f"{unscheduled_patient}, {eq}", patient_info)
                            caregivers[caregiver] = new_patient_info

                    # Remove the patient from patients_unscheduled as they are now scheduled
                    patients_unscheduled[unscheduled_patient] = [(eq, count - 1) for eq, count in equipment_list if
                                                                 count > 1]
                    if not any(count > 0 for _, count in patients_unscheduled[unscheduled_patient]):
                        del patients_unscheduled[unscheduled_patient]

                    # Remove the unavailable patient from the unavailability list
                    del patients_unavailability_list[unavailable_patient]

                    replacement_found = True
                    break
            if replacement_found:
                break

    save_schedule_to_excel(schedule, input_file_path, output_file_path)

def find_caregivers_for_patients(schedule, patients_unavailability_list):
    # Create an empty list to store caregivers
    caregivers_list = []

    # Iterate over each patient in unavailability list
    for patient, unavailable_time in patients_unavailability_list.items():
        if unavailable_time in schedule:
            caregivers = schedule[unavailable_time]  # Get caregivers at that specific time

            # Iterate through each caregiver
            for caregiver, patient_info in caregivers.items():
                # Parse the patient info to find patient-equipment pairings
                matches = re.findall(r"(Patient\d+), (Equipment\d+)", patient_info)
                for match_patient, match_equipment in matches:
                    if match_patient == patient:
                        # Append a tuple with caregiver, unavailable_time, and patient
                        caregivers_list.append((caregiver, unavailable_time, patient))
                        break

    return caregivers_list

def create_a_list_of_patients_and_their_equipment_in_caregiver_unavailable_slot(schedule, caregiver_updated_unavailability):
    patients_list_in_unavailable_slot = []
    for caregiver, unavailability_time in caregiver_updated_unavailability.items():
        if unavailability_time in schedule:
            caregiver_schedule = schedule[unavailability_time].get(caregiver, '')
            if caregiver_schedule:
                # Split the schedule string by ', ' to get patients and equipment pairs
                # Use '; ' to split different patient entries
                patient_entries = caregiver_schedule.split(', ')

                # Initialize a dictionary to store patient data
                patient_data = {}

                # Iterate over entries in pairs
                for i in range(0, len(patient_entries), 2):
                    if i + 1 < len(patient_entries):  # Ensure there is a pair
                        patient = patient_entries[i].strip()
                        equipment = patient_entries[i + 1].strip()

                        if patient and equipment:
                            if patient not in patient_data:
                                patient_data[patient] = []
                            patient_data[patient].append(equipment)

                # Add formatted data to the result list
                for patient, equipments in patient_data.items():
                    patients_list_in_unavailable_slot.append({patient: equipments})
    # Print the result
    print("Patients List in Unavailable Slot:")
    for item in patients_list_in_unavailable_slot:
        for patient, equipments in item.items():
            print(f"{patient}: {equipments}")

    return patients_list_in_unavailable_slot

def create_schedule_from_excel(schedule_file_path):
    # Read the Excel file
    df = pd.read_excel(schedule_file_path, engine='openpyxl')

    # Initialize the schedule dictionary
    schedule = {}

    # Iterate through each row in the dataframe
    for _, row in df.iterrows():
        # Extract time from the row (assuming the time is in the index)
        time = row.name  # Use row index for time

        # Initialize caregivers dictionary
        caregivers = {}

        # Iterate through each caregiver column
        for caregiver in df.columns:
            # Skip the index column
            if caregiver == 'Time':
                continue
            if pd.notna(row[caregiver]):  # Check if the cell is not NaN
                caregivers[caregiver] = row[caregiver]

        # Add to the schedule dictionary
        schedule[time] = caregivers

    return schedule

def reformat_schedule_with_regex(original):
    schedule = {}

    # Regex pattern to match time (HH:MM) format
    time_pattern = re.compile(r'^\d{1,2}:\d{2}$')

    for key, value in original.items():
        # Check if 'Unnamed: 0' contains a time entry
        time_entry = value.get('Unnamed: 0', '')
        if time_pattern.match(time_entry):
            # Extract hour from time entry
            hour = int(time_entry.split(':')[0])

            # Initialize a dictionary to store caregiver data
            caregivers = {}

            # Iterate over the key-value pairs in the value dictionary
            for sub_key, sub_value in value.items():
                # Skip the time entry key
                if sub_key == 'Unnamed: 0':
                    continue

                # Add to the caregivers dictionary
                caregivers[sub_key] = sub_value

            # Add to the new schedule
            schedule[hour] = caregivers

    return schedule

def main():
    input_file = r"C:\Users\morsh\Desktop\personal_projects\soroka_solution\small_rehabilitation_data.xlsx"
    output_file = r"C:\Users\morsh\Desktop\personal_projects\soroka_solution\updated_schedule_after_changing_unavailability.xlsx"
    schedule_file_path = r"C:\Users\morsh\Desktop\personal_projects\soroka_solution\generated_schedule.xlsx"
    schedule = create_schedule_from_excel(schedule_file_path)
    updated_schedule = reformat_schedule_with_regex(schedule)

    patients_unavailability_list = {'Patient3': 15}
    replace_unavailable_patient_with_matching_available_one(input_file, schedule_file_path, patients_unavailability_list, output_file)

if __name__ == "__main__":
    main()

