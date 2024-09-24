import pandas as pd


def parse_time_slot(time_range):
    start_time, end_time = map(int, time_range.split('-'))
    start_hour = start_time // 100
    end_hour = end_time // 100
    # Ensure the range includes the start hour and goes up to the hour before the next start
    return list(range(end_hour, start_hour + 1))



def get_students_schedule(file_path, crn):
    # Load the dataset
    df = pd.read_csv(file_path)
    
    # Check if the CRN exists in the dataset
    if crn not in df['ر.مرجعي'].unique():
        return None  # Return None if CRN is not found
    
    # Find all student IDs associated with the given CRN
    students_with_crn = df[df['ر.مرجعي'] == crn]['الرقم التدريبي'].unique()
    
    # Filter the DataFrame for these students (to get all their courses)
    df = df[df['الرقم التدريبي'].isin(students_with_crn)]
    
    # Group data by student ID and aggregate their time slots
    student_schedules = {}
    for _, row in df.iterrows():
        student_id = row['الرقم التدريبي']
        day = row['اليوم'].strip()  # Ensure to strip any whitespace from day names
        time_slot = parse_time_slot(row['الوقت'])
        if student_id not in student_schedules:
            student_schedules[student_id] = {}
        if day not in student_schedules[student_id]:
            student_schedules[student_id][day] = []
        student_schedules[student_id][day].extend(time_slot)
        student_schedules[student_id][day] = sorted(set(student_schedules[student_id][day]))  # Sort and remove duplicates
    
    return student_schedules

def find_common_free_slots(students_schedules):
    
    if not students_schedules:
        return "No schedules found for the entered CRN. Please check the CRN and try again."
    
    # Initialize a dictionary to hold the busy slots for each day
    busy_slots = {day: set() for day in ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس']}
    
    # Aggregate all busy slots for each student into a single set for each day
    for schedule in students_schedules.values():
        for day, slots in schedule.items():
            busy_slots[day].update(slots)
    
    # Determine all possible slots in a day (assuming slots from 8AM to 8PM)
    all_possible_slots = set(range(8, 16))  # 8AM to 8PM
    
    # Determine free slots by subtracting busy slots from all possible slots
    free_slots = {day: sorted(all_possible_slots - slots) for day, slots in busy_slots.items()}
    
    return free_slots


