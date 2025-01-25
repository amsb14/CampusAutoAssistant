import pandas as pd



# Function to calculate past four terms based on the term input
def calculate_past_terms(term):
    new_term = str(term)[1:5]
    year = int(new_term[:3])
    half = int(new_term[3])
    past_terms = []

    for _ in range(4):
        if half == 2:
            half = 1
        else:
            half = 2
            year -= 1
        past_terms.append(f"{year}{half}")

    return new_term, past_terms


# Function to extract the term level based on input term
# def determine_term_level(current_term, training_number):
#     new_term, past_terms = calculate_past_terms(current_term)
#     prefix = str(training_number)[:4]
#     if prefix == new_term:
#         return 1
#     elif prefix in past_terms:
#         return past_terms.index(prefix) + 2
#     else:
#         return "Unknown"

# Function to extract the term level based on input term (uncomment above code and comment this if necessary as current code may have mismatching problem in defining terms) Also check line 113 lambda row: determine_term_level
def determine_term_level(training_number):
    # Calculate past four terms

    prefix = str(training_number)[:4]
    if prefix == "4452":
        return 1
    elif prefix == "4451":
        return 2
    elif prefix == "4443":
        return 3
    elif prefix == "4442":
        return 4
    elif prefix == "4441":
        return 5
    else:
        return "متدربين متجاوزين الفصول المحددة"
    
    
# Function to check compliance
def check_compliance(student_level, course_level):
    if (
        (student_level == 1 and 1 <= course_level <= 4)
        or (student_level == 2 and 2 <= course_level <= 4)
        or (student_level == 3 and 3 <= course_level <= 4)
        or (student_level == 4 and course_level == 4)
    ):
        return "متوافق"
    return "غير متوافق"


# Function to determine commitment
def determine_commitment(compliance_column):
    return "غير ملتزم" if "غير متوافق" in compliance_column.values else "ملتزم"


# Main analysis function
def analyze_compliance(uploaded_file, current_term, course_mapping):
    ss05_data = pd.read_csv(uploaded_file)

    # Clean and map data
    ss05_data['عنوان المقرر'] = ss05_data['عنوان المقرر'].str.strip()
    ss05_data['Course Level'] = ss05_data['عنوان المقرر'].map(course_mapping)
    ss05_data['Student Level'] = ss05_data.apply(
        # lambda row: determine_term_level(current_term, row['الرقم التدريبي']),
        lambda row: determine_term_level(row['الرقم التدريبي']),
        axis=1,
    )

    # Apply compliance check
    ss05_data['Compliance'] = ss05_data.apply(
        lambda row: check_compliance(row['Student Level'], row['Course Level']), axis=1
    )

    # Group by student and calculate commitment
    ss05_data['Commitment'] = ss05_data.groupby('الرقم التدريبي')['Compliance'].transform(determine_commitment)

    # Summarize data
    commitment_summary = (
        ss05_data.groupby(['Student Level', 'Commitment'])['الرقم التدريبي']
        .nunique()
        .reset_index()
        .rename(columns={'الرقم التدريبي': 'Unique Student Count'})
    )

    # Pivot and calculate statistics
    result_table = (
        commitment_summary.pivot(index='Student Level', columns='Commitment', values='Unique Student Count')
        .fillna(0)
        .rename(columns={'ملتزم': 'عدد الملتزمين', 'غير ملتزم': 'الغير ملتزمين'})
        .reset_index()
    )

    # Add missing levels and sort
    all_levels = ['المستوى الأول', 'المستوى الثاني', 'المستوى الثالث', 'المستوى الرابع', 'المستوى الخامس', 'متدربين متجاوزين الفصول المحددة']
    result_table['Student Level'] = result_table['Student Level'].replace({
        1: 'المستوى الأول',
        2: 'المستوى الثاني',
        3: 'المستوى الثالث',
        4: 'المستوى الرابع',
        5: 'المستوى الخامس',
    })
    for level in all_levels:
        if level not in result_table['Student Level'].values:
            result_table = pd.concat([result_table, pd.DataFrame({'Student Level': [level], 'عدد الملتزمين': [0], 'الغير ملتزمين': [0]})])
    result_table = result_table.set_index('Student Level').loc[all_levels].reset_index()

    # Calculate totals and percentages
    result_table['المجموع'] = result_table['عدد الملتزمين'] + result_table['الغير ملتزمين']
    result_table['نسبة الملتزمين'] = (result_table['عدد الملتزمين'] / result_table['المجموع'] * 100).fillna(0).round(2).astype(str) + '%'
    result_table['نسبة الغير ملتزمين'] = (result_table['الغير ملتزمين'] / result_table['المجموع'] * 100).fillna(0).round(2).astype(str) + '%'

    # Add overall summary
    total_compliant = result_table['عدد الملتزمين'].sum()
    total_non_compliant = result_table['الغير ملتزمين'].sum()
    total_students = result_table['المجموع'].sum()
    overall_summary = pd.DataFrame({
        'Student Level': ['الإجمالي'],
        'عدد الملتزمين': [total_compliant],
        'الغير ملتزمين': [total_non_compliant],
        'المجموع': [total_students],
        'نسبة الملتزمين': [(total_compliant / total_students * 100) if total_students > 0 else 0],
        'نسبة الغير ملتزمين': [(total_non_compliant / total_students * 100) if total_students > 0 else 0],
    })
    overall_summary['نسبة الملتزمين'] = overall_summary['نسبة الملتزمين'].round(2).astype(str) + '%'
    overall_summary['نسبة الغير ملتزمين'] = overall_summary['نسبة الغير ملتزمين'].round(2).astype(str) + '%'
    result_table = pd.concat([result_table, overall_summary], ignore_index=True)
    
    # Rename the "Student Level" column for display
    result_table = result_table.rename(columns={'Student Level': 'المستوى التدريبي'})

    return result_table
